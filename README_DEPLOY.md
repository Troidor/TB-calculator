# Deploying to Render - Secure Calculator (Yellow inputs in Calcs!C)

1. Create a GitHub repo and push the project files.

2. Put your Excel workbook into the `excel/` folder. Name it `TB_v18.xlsx`.

3. On Render: New -> Web Service -> Connect GitHub -> select repo and branch.

4. Settings:
   - Environment: Python 3
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app`

5. Environment variables on Render:
   - `APP_PASSWORD` = your password
   - `SECRET_KEY` = a long random secret

Notes:
- Inputs detection: the app treats yellow-shaded cells in column C of sheet 'Calcs' as inputs.
- Outputs: all formula cells in sheet 'Calcs' are evaluated and displayed.
- Do not commit secrets to GitHub.
