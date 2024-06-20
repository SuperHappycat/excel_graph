import excel_chart

# Example data
data = {
    'Month': ['January', 'February', 'March', 'April', 'May', 'June', 'July'],
    'Sales': [150, 200, 250, 300, 350, 400, 450]
}

# Create Excel file with chart
excel_chart.create_excel_with_chart(data, chart_title='Monthly Sales Data', x_label='Month', y_label='Sales', filename='sales_data_with_chart.xlsx')