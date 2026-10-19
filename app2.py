from flask import Flask, request, render_template_string
import pandas as pd
import requests
import base64
import time
import os
import logging

app = Flask(__name__)

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s [%(levelname)s] %(message)s')

# API endpoints
FILE_API_MAP = {
    'financial_file': 'http://10.225.9.180:32085/v1/financialtransaction',
    'non_financial_file': 'http://10.225.9.180:32085/v1/nonfinancialtransaction',
    'card_file': 'http://10.225.9.180:32085/v1/cardtransaction'
}

# HTML template with CSS styling and message placeholder
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html>
<head>
    <title>SAS Transaction Upload Portal</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f4faff;
            margin: 0; padding: 0;
        }
        .container {
            max-width: 700px;
            margin: 50px auto;
            background-color: #fff;
            border-radius: 8px;
            box-shadow: 0 0 15px rgba(0,0,50,0.1);
            padding: 30px;
        }
        h2 {
            color: #005A9C;
            text-align: center;
        }
        form label {
            font-weight: bold;
            color: #333;
        }
        input[type="text"], input[type="password"], input[type="file"] {
            width: 100%;
            padding: 10px;
            margin-top: 5px;
            margin-bottom: 20px;
            border: 1px solid #ccc;
            border-radius: 4px;
            box-sizing: border-box;
        }
        input[type="submit"] {
            background-color: #005A9C;
            color: white;
            border: none;
            padding: 12px 20px;
            font-size: 16px;
            cursor: pointer;
            width: 100%;
            border-radius: 4px;
            transition: background-color 0.3s ease;
        }
        input[type="submit"]:hover {
            background-color: #003d6a;
        }
        .message {
            margin-top: 20px;
            padding: 15px;
            border-radius: 6px;
            font-size: 16px;
        }
        .message.success {
            background-color: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }
        .message.error {
            background-color: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
            font-size: 14px;
        }
        table, th, td {
            border: 1px solid #ddd;
        }
        th {
            background-color: #005A9C;
            color: white;
            padding: 8px;
            text-align: center;
        }
        td {
            padding: 8px;
            text-align: center;
        }
    </style>
</head>
<body>
    <div class="container">
        <h2>Upload Transaction Excel Files</h2>
        <form action="/" method="post" enctype="multipart/form-data">
            <label>Username:</label><br>
            <input type="text" name="username" required><br>

            <label>Password:</label><br>
            <input type="password" name="password" required><br>

            <label>Upload Financial File:</label><br>
            <input type="file" name="financial_file" accept=".xlsx"><br>

            <label>Upload Non-Financial File:</label><br>
            <input type="file" name="non_financial_file" accept=".xlsx"><br>

            <label>Upload Card File:</label><br>
            <input type="file" name="card_file" accept=".xlsx"><br>

            <input type="submit" value="Upload & Send">
        </form>

        <div class="message">
            {{ messages|safe }}
        </div>
    </div>
</body>
</html>
'''

def nest_keys(flat_dict):
    nested = {}
    numeric_keys = {
        'CustomerAccount.customerAccountBalanceLocalCurrency',
        'Transaction.transactionAmountLocalCurrency'
    }
    for key, value in flat_dict.items():
        if key in numeric_keys:
            # Convert to number (float), empty string becomes None
            try:
                value = float(value) if value != '' else None
            except ValueError:
                value = None
        else:
            value = str(value)

        if '.' in key:
            parts = key.split('.')
            d = nested
            for part in parts[:-1]:
                d = d.setdefault(part, {})
            d[parts[-1]] = value
        else:
            nested[key] = value
    return nested

@app.route('/', methods=['GET', 'POST'])
def upload_excel():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')

        if not username or not password:
            return render_template_string(
                HTML_TEMPLATE,
                messages="<p class='error'>Username and password are required.</p>"
            )

        basic_token = base64.b64encode(f"{username}:{password}".encode()).decode()
        headers = {
            "Authorization": f"Basic {basic_token}",
            "Content-Type": "application/json"
        }

        all_jsons = {}
        all_responses = []
        processed_files = 0

        for file_field, api_url in FILE_API_MAP.items():
            uploaded_file = request.files.get(file_field)
            if uploaded_file:
                filename = f'temp_{file_field}.xlsx'
                uploaded_file.save(filename)

                try:
                    df = pd.read_excel(filename)
                    if df.empty:
                        app.logger.warning(f"Uploaded file for {file_field} is empty.")
                        continue

                    nested_json_list = []
                    for _, row in df.iterrows():
                        flat_data = row.fillna('').to_dict()
                        nested_json = nest_keys(flat_data)
                        nested_json_list.append(nested_json)

                    all_jsons[file_field] = {
                        'api_url': api_url,
                        'data': nested_json_list
                    }
                    processed_files += 1
                except Exception as e:
                    app.logger.error(f"Error reading {file_field}: {e}")
                finally:
                    if os.path.exists(filename):
                        os.remove(filename)

        for file_field, api_info in all_jsons.items():
            api_url = api_info['api_url']
            json_list = api_info['data']

            for row_json in json_list:
                try:
                    app.logger.info(f"Sending to {api_url}: {row_json}")
                    response = requests.post(api_url, json=row_json, headers=headers)
                    app.logger.info(f"Received status: {response.status_code}")
                    app.logger.info(f"Response content: {response.text}")

                    if response.status_code == 401:
                        return render_template_string(
                            HTML_TEMPLATE,
                            messages="<p class='error'>? Invalid username or password.</p>"
                        )

                    try:
                        res_data = response.json()
                        res_data['source'] = file_field
                        res_data['status_code'] = response.status_code
                        all_responses.append(res_data)
                    except Exception as json_err:
                        app.logger.warning(f"JSON parsing failed: {json_err}")
                        all_responses.append({
                            'source': file_field,
                            'status_code': response.status_code,
                            'response': response.text
                        })

                except Exception as api_err:
                    app.logger.error(f"Error sending to API: {api_err}")
                    all_responses.append({
                        'source': file_field,
                        'status_code': 'error',
                        'response': str(api_err)
                    })

        if processed_files == 0:
            return render_template_string(
                HTML_TEMPLATE,
                messages="<p class='error'>No valid Excel file uploaded.</p>"
            )

        if all_responses:
            df_all = pd.json_normalize(all_responses)
            app.logger.info(f"Normalized response columns: {list(df_all.columns)}")

            selected_columns = [
                'customerID',
                'transactionID',
                'fraudScore',
                'recommendedAction'
            ]
            rename_map = {
                'customerID': 'CustomerID',
                'transactionID': 'TransactionID',
                'fraudScore': 'Fraud_Score',
                'recommendedAction': 'Recommended_Action'
            }

            available_columns = [col for col in selected_columns if col in df_all.columns]

            if available_columns:
                df_filtered = df_all[available_columns].rename(columns=rename_map)
                response_table = df_filtered.to_html(index=False, border=1, justify="center")
            else:
                response_table = "<p>No matching response fields found.</p>"
        else:
            response_table = "<p>No responses received.</p>"

        return render_template_string(
            HTML_TEMPLATE,
            messages=f"<p class='success'>{processed_files} file(s) processed successfully.</p><br><h3>API Responses:</h3>{response_table}"
        )

    return render_template_string(HTML_TEMPLATE, messages="")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5055, debug=True)
