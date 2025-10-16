import pandas as pd
import numpy as np
import sqlite3 
df = pd.read_csv('AppleStore.csv')
print(df)
df = df.dropna()

conn = sqlite3.connect(":memory:")
df.to_sql("applestore", conn, index=False, if_exists="replace")

Query_1 = """
select 
     prime_genre,
     sum(case 
     when price = 0 then 1 else 0 end )
       as Free_Apps,
     sum(case 
     when price > 0 then 1 else 0 end )
       as Paid_Apps,
     Count(*) as total_apps
     from applestore
     group by prime_genre;

"""

result = pd.read_sql_query(Query_1, conn)
print(result)



Paid_Query_rating = """
select 
     prime_genre,
     sum(case when user_rating > 4 then  1 else 0 end ) as Rating_above_4,
     sum(case when user_rating >= 4 and price > 0 then 1 else 0 end ) as Paid_Rating_Above_4,
     Sum(case when user_rating >= 4 and price = 0 then 1 else 0 end ) as Free_Rating_above_4,
     sum(case when user_rating < 3 then 1 else 0 end ) as Worst_Ratings
     from applestore
     group by prime_genre
     order by Worst_Ratings DESC;
"""

result2 = pd.read_sql_query(Paid_Query_rating, conn)
print(result2)


Avg_size = """
SELECT 
    prime_genre,
    ROUND(AVG(CASE WHEN price = 0 THEN size_bytes / 1024 / 1024 END), 2) AS Avg_Size_MB_Free,
    ROUND(AVG(CASE WHEN price > 0 THEN size_bytes / 1024 / 1024 END), 2) AS Avg_Size_MB_Paid,
    SUM(CASE WHEN price = 0 THEN rating_count_tot ELSE 0 END) AS Free_Reviews,
    SUM(CASE WHEN price > 0 THEN rating_count_tot ELSE 0 END) AS Paid_Reviews,
    SUM(rating_count_tot) AS Total_Reviews,
    ROUND(AVG(user_rating), 2) AS Avg_Rating
FROM applestore
GROUP BY prime_genre
ORDER BY Avg_rating DESC ;
"""

result3 = pd.read_sql_query(Avg_size, conn)
print(result3)


Compare = """
SELECT 
    ROUND(AVG(CASE WHEN price = 0 THEN user_rating END), 2) AS Avg_Rating_Free,
    ROUND(AVG(CASE WHEN price > 0 THEN user_rating END), 2) AS Avg_Rating_Paid
FROM applestore;
"""
result4 = pd.read_sql_query(Compare, conn)
print(result4)


output_file = "AppleDashBoard_Data.xlsx"

with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    # Each query result in a separate sheet
    result.to_excel(writer, sheet_name="Genre", index=False)
    result2.to_excel(writer, sheet_name="Rating", index=False)
    result3.to_excel(writer, sheet_name="Reviews_Size", index=False)
    result4.to_excel(writer, sheet_name="Compare", index=False)

print(f"All data written successfully to {output_file}")