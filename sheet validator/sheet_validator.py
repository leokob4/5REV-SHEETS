def validate_sheets():
    try:
        wb = openpyxl.load_workbook("user_sheets/db.xlsx")
        expected = {
            "users": ["id", "username", "password_hash", "role"],
            "tools": ["id", "name", "description", "path"],
            "access": ["role", "tools"]
        }

        results = {}
        for sheet_name, expected_cols in expected.items():
            if sheet_name not in wb.sheetnames:
                results[sheet_name] = "❌ Missing Sheet"
                continue

            sheet = wb[sheet_name]
            headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
            if headers != expected_cols:
                results[sheet_name] = f"⚠️ Invalid Headers: {headers}"
            else:
                results[sheet_name] = "✅ OK"

        return results
    except Exception as e:
        return {"global": f"❌ {str(e)}"}
