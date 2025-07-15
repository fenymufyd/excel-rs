import py_excel_rs
from datetime import datetime, timezone, timedelta
import pandas as pd

naive_date = datetime.now()
tz_date = datetime.now(timezone(timedelta(hours=8)))

data = [[naive_date, tz_date, "hello", 10, 10.888]]
df = pd.DataFrame(data, columns=["Date", "Timezone Date", "hi", "number1", "float2"])

xlsx = py_excel_rs.df_to_xlsx(df, should_infer_types=False)

with open('report.xlsx', 'wb') as f:
    f.write(xlsx)