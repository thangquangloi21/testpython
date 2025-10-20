from sql_conn import SQLQuery

db = SQLQuery(
        server='10.239.1.54',
        database='khsx_log',
        username='sa',
        password='123456',
        driver='{SQL Server}'
    )

try:
        # Ví dụ truy vấn SELECT
        query_select = "SELECT * FROM Employees"
        # results = db.execute_query(query_select)
        # if results:
        #     for row in results:
        #         print(f'ID: {row.ID}, Name: {row.Name}, Age: {row.Age}, Department: {row.Department}')
        
        # Ví dụ truy vấn INSERT với tham số
        query_insert = "INSERT INTO Employees (Name, Age, Department) VALUES (?, ?, ?)"
        db.execute_query(query_insert, ('Nguyen Van B', 30, 'IT'))
        
        # # Ví dụ truy vấn UPDATE với tham số
        # query_update = "UPDATE Employees SET Age = ? WHERE Name = ?"
        # db.execute_query(query_update, (31, 'Nguyen Van A'))
        
        # # # Ví dụ truy vấn DELETE với tham số
        # query_delete = "DELETE FROM Employees WHERE Name = ?"
        # db.execute_query(query_delete, ('Nguyen Van A',))
        
        # Truy vấn lại để kiểm tra
        results = db.execute_query(query_select)
        if results:
            for row in results:
                print(f'ID: {row.ID}, Name: {row.Name}, Age: {row.Age}, Department: {row.Department}')
        
finally:
        # Đóng kết nối
        db.close()