use std::{
    collections::VecDeque,
    io::{IoSlice, Seek, Write},
};

use anyhow::Result;
use zip::{write::SimpleFileOptions, ZipWriter};

pub struct Sheet<'a, W: Write + Seek> {
    pub sheet_buf: &'a mut ZipWriter<W>,
    pub _name: String,
    // pub id: u16,
    // pub is_closed: bool,
    current_row_num: u32,
    global_shared_vec: Vec<u8>,
    global_shared_letter_vec: Vec<u8>,
}

fn col_to_letter(vec: &mut Vec<u8>, col: usize) -> &[u8] {
    // let mut result = Vec::with_capacity(2);
    let mut col = col as i16;

    loop {
        vec.push(b'A' + (col % 26) as u8);
        col = col / 26 - 1;
        if col < 0 {
            break;
        }
    }

    vec.reverse();
    vec
}

fn ref_id(vec: &mut Vec<u8>, col: usize, row: ([u8; 9], usize)) -> Result<([u8; 12], usize)> {
    let mut final_arr: [u8; 12] = [0; 12];
    let letter = col_to_letter(vec, col);

    let mut pos: usize = 0;
    for c in letter {
        final_arr[pos] = *c;
        pos += 1;
    }

    let (row_in_chars_arr, digits) = row;

    for i in 0..digits {
        final_arr[pos] = row_in_chars_arr[(8 - digits) + i + 1];
        pos += 1;
    }

    vec.clear();

    Ok((final_arr, pos))
}

fn num_to_bytes(n: u32) -> ([u8; 9], usize) {
    // Convert from number to string manually
    let mut row_in_chars_arr: [u8; 9] = [0; 9];
    let mut row = n;
    let mut char_pos = 8;
    let mut digits = 0;
    while row > 0 {
        row_in_chars_arr[char_pos] = b'0' + (row % 10) as u8;
        row = row / 10;
        char_pos -= 1;
        digits += 1;
    }

    (row_in_chars_arr, digits)
}

fn escape_in_place(bytes: &[u8]) -> (VecDeque<&[u8]>, VecDeque<usize>) {
    let mut special_chars: VecDeque<&[u8]> = VecDeque::new();
    let mut special_char_pos: VecDeque<usize> = VecDeque::new();
    let len = bytes.len();
    for x in 0..len {
        let _ = match bytes[x] {
            b'<' => {
                special_chars.push_back(b"&lt;".as_slice());
                special_char_pos.push_back(x);
            }
            b'>' => {
                special_chars.push_back(b"&gt;".as_slice());
                special_char_pos.push_back(x);
            }
            b'\'' => {
                special_chars.push_back(b"&apos;".as_slice());
                special_char_pos.push_back(x);
            }
            b'&' => {
                special_chars.push_back(b"&amp;".as_slice());
                special_char_pos.push_back(x);
            }
            b'"' => {
                special_chars.push_back(b"&quot;".as_slice());
                special_char_pos.push_back(x);
            }
            _ => (),
        };
    }

    (special_chars, special_char_pos)
}

impl<'a, W: Write + Seek> Sheet<'a, W> {
    pub fn new(name: String, id: u16, writer: &'a mut ZipWriter<W>) -> Result<Self> {
        let options = SimpleFileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .compression_level(Some(1))
            .large_file(true);

        writer
            .start_file(format!("xl/worksheets/sheet{}.xml", id), options)?;

        // Writes Sheet Header
        writer.write(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheetData>").ok();

        Ok(Sheet {
            sheet_buf: writer,
            // id,
            _name: name,
            // is_closed: false,
            // col_num_to_letter: Vec::with_capacity(64),
            current_row_num: 0,
            global_shared_vec: Vec::new(),
            global_shared_letter_vec: Vec::new(),
        })
    }

    // TOOD: Use ShortVec over Vec for cell ID
    pub fn write_row(&mut self, data: Vec<&[u8]>) -> Result<()> {
        self.current_row_num += 1;

        // let final_vec = &mut self.global_shared_vec;

        // TODO: Proper Error Handling
        let (row_in_chars_arr, digits) = num_to_bytes(self.current_row_num);

        // self.global_shared_vec.write(b"<row r=\"")?;
        // self.global_shared_vec
        //     .write(&row_in_chars_arr[9 - digits..])?;
        // self.global_shared_vec.write(b"\">")?;

        self.global_shared_vec.write_vectored(&[
            IoSlice::new(b"<row r=\""),
            IoSlice::new(&row_in_chars_arr[9 - digits..]),
            IoSlice::new(b"\">"),
        ])?;

        let mut col = 0;
        for datum in data {
            let (ref_id, pos) = ref_id(
                &mut self.global_shared_letter_vec,
                col,
                (row_in_chars_arr, digits),
            )?;

            // self.global_shared_vec.write(b"<c r=\"")?;
            // self.global_shared_vec.write(&ref_id.as_slice()[0..pos])?;
            // self.global_shared_vec.write(b"\" t=\"str\"><v>")?;

            self.global_shared_vec.write_vectored(&[
                IoSlice::new(b"<c r=\""),
                IoSlice::new(&ref_id.as_slice()[0..pos]),
                IoSlice::new(b"\" t=\"str\"><v>"),
            ])?;

            let (mut chars, chars_pos) = escape_in_place(datum);
            let mut current_pos = 0;
            for char_pos in chars_pos {
                self.global_shared_vec.write_vectored(&[
                    IoSlice::new(&datum[current_pos..char_pos]),
                    IoSlice::new(&chars.pop_front().unwrap()),
                ])?;
                current_pos = char_pos + 1;
            }

            self.global_shared_vec.write_vectored(&[
                IoSlice::new(&datum[current_pos..]),
                IoSlice::new(b"</v></c>"),
            ])?;

            col += 1;
        }

        self.global_shared_vec.write(b"</row>")?;

        if self.current_row_num % 100_000 == 0 {
            self.flush()?;
        }

        Ok(())
    }

    fn flush(&mut self) -> Result<()> {
        self.sheet_buf.write_all(&self.global_shared_vec)?;
        self.global_shared_vec.clear();
        Ok(())
    }

    pub fn close(&mut self) -> Result<()> {
        self.flush()?;
        self.sheet_buf.write(b"</sheetData></worksheet>")?;
        Ok(())
    }
}
