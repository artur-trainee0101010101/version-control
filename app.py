from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from io import BytesIO

app = Flask(__name__)

def read_excel_data(file):
    """Read Excel file and return a DataFrame."""
    return pd.read_excel(file)

def process_data(df):
    """Process DataFrame into a simple, flat structure."""
    processed_data = {}
    processed_data['total_sales'] = df['Revenue'].sum()
    processed_data['top_product'] = df.loc[df['Revenue'].idxmax(), 'Product']
    processed_data['sales_by_market'] = df.groupby('Market')['Revenue'].sum().to_dict()
    processed_data['num_markets'] = df['Market'].nunique()
    return processed_data

def generate_strategy_text(data):
    """Generate strategy text based on processed data."""
    text = "Strategic Analysis Report\n\n"
    text += f"Total Sales: ${data['total_sales']:,.2f}\n"
    text += f"Top Performing Product: {data['top_product']}\n"
    text += f"Number of Markets: {data['num_markets']}\n\n"
   
    text += "Sales by Market:\n"
    for market, sales in data['sales_by_market'].items():
        text += f"- {market}: ${sales:,.2f}\n"
   
    text += "\nStrategic Recommendations:\n"
    text += f"1. Focus on expanding market share for {data['top_product']}.\n"
    if data['num_markets'] < 4:
        text += "2. Explore opportunities to enter new markets.\n"
    else:
        text += "2. Optimize operations in existing markets to improve efficiency.\n"
    if max(data['sales_by_market'].values()) / min(data['sales_by_market'].values()) > 2:
        text += "3. Investigate and address performance discrepancies between markets.\n"
    text += "4. Conduct customer surveys to identify product improvement opportunities.\n"
    return text

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files['file']
        if file and file.filename.endswith('.xlsx'):
            df = read_excel_data(file)
            processed_data = process_data(df)
            strategy_text = generate_strategy_text(processed_data)
            return render_template('result.html', strategy=strategy_text)
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)