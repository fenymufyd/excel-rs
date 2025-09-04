use super::format::XlsxFormatter;
use anyhow::Result;
use std::{io::Seek, io::Write};
use zip::ZipWriter;

use super::sheet::Sheet;
use super::typed_sheet::TypedSheet;

pub struct WorkBook<W: Write + Seek> {
    formatter: XlsxFormatter<W>,
    num_of_sheets: u16,
}

impl<W: Write + Seek> WorkBook<W> {
    pub fn new(writer: W) -> Self {
        let zip_writer = ZipWriter::new(writer);

        WorkBook {
            formatter: XlsxFormatter::new(zip_writer),
            num_of_sheets: 0,
        }
    }

    pub fn get_worksheet(&'_ mut self, name: String) -> Result<Sheet<'_, W>> {
        self.num_of_sheets += 1;
        Sheet::new(name, self.num_of_sheets, &mut self.formatter.zip_writer)
    }

    pub fn get_typed_worksheet(&'_ mut self, name: String) -> Result<TypedSheet<'_, W>> {
        self.num_of_sheets += 1;
        TypedSheet::new(name, self.num_of_sheets, &mut self.formatter.zip_writer)
    }

    pub fn finish(self) -> Result<W> {
        let result = self.formatter.finish(self.num_of_sheets)?;
        Ok(result)
    }
}
