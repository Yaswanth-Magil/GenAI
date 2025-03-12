import pandas as pd
from sqlalchemy import create_engine

# Read the Excel file (make sure openpyxl is installed for .xlsx files)
df = pd.read_excel('/Users/yash/Downloads/Today/Splitted/GenAI/Previous month data.xlsx', engine='openpyxl')

# Create a SQLAlchemy engine for a local MySQL database.
# Replace 'username', 'password', 'localhost', and 'database_name' with your details.
engine = create_engine('mysql+pymysql://root:Yaswanth123.@localhost/genai')

# Import the DataFrame into MySQL. 
# The 'if_exists' parameter can be set to 'fail', 'replace', or 'append'
df.to_sql(name='reviews_trend_dummy', con=engine, if_exists='append', index=False)