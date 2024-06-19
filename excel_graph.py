import pandas as pd

def create_dataframe(data):
    """
    Creates a DataFrame from a dictionary of data.
    
    Parameters:
    - data: A dictionary where keys are column names and values are lists of data.
    
    Returns:
    - A pandas DataFrame.
    """
    return pd.DataFrame(data)

def write_dataframe_to_excel(df, filename, sheet_name='Sheet1'):
    """
    Writes a DataFrame to an Excel file.
    
    Parameters:
    - df: The DataFrame to write.
    - filename: The name of the output Excel file.
    - sheet_name: The sheet name in the Excel file (default is 'Sheet1').
    
    Returns:
    - A tuple containing the workbook and worksheet objects.
    """
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        return writer.book, writer.sheets[sheet_name]

def add_chart_to_excel(workbook, worksheet, categories_range, values_range, chart_title, x_label, y_label):
    """
    Adds a chart to an Excel worksheet.
    
    Parameters:
    - workbook: The Excel workbook object.
    - worksheet: The Excel worksheet object.
    - categories_range: The range for the chart categories (x-axis).
    - values_range: The range for the chart values (y-axis).
    - chart_title: The title of the chart.
    - x_label: The label for the x-axis.
    - y_label: The label for the y-axis.
    """
    # Create a chart object
    chart = workbook.add_chart({'type': 'line'})
    
    # Configure the series of the chart
    chart.add_series({
        'name': y_label,
        'categories': categories_range,
        'values': values_range,
        'line': {'color': 'blue'},
        'marker': {'type': 'circle', 'size': 6, 'border': {'color': 'blue'}, 'fill': {'color': 'yellow'}}
    })
    
    # Add chart title and axis labels
    chart.set_title({'name': chart_title})
    chart.set_x_axis({'name': x_label, 'name_font': {'size': 14, 'bold': True}, 'num_font': {'italic': True}})
    chart.set_y_axis({'name': y_label, 'name_font': {'size': 14, 'bold': True}, 'num_font': {'italic': True}})
    
    # Add legend
    chart.set_legend({'position': 'bottom'})
    
    # Set a chart style
    chart.set_style(10)
    
    # Insert the chart into the worksheet
    worksheet.insert_chart('D2', chart)

def create_excel_with_chart(data, chart_title, x_label, y_label, filename='chart.xlsx'):
    """
    Creates an Excel file with a chart based on the provided data.
    
    Parameters:
    - data: A dictionary where keys are column names and values are lists of data.
    - chart_title: The title of the chart.
    - x_label: The label for the x-axis.
    - y_label: The label for the y-axis.
    - filename: The name of the output Excel file (default is 'chart.xlsx').
    """
    df = create_dataframe(data)
    workbook, worksheet = write_dataframe_to_excel(df, filename)
    
    # Define the categories and values range
    categories_range = [worksheet.name, 1, 0, len(df), 0]
    values_range = [worksheet.name, 1, 1, len(df), 1]
    
    add_chart_to_excel(workbook, worksheet, categories_range, values_range, chart_title, x_label, y_label)

