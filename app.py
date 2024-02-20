from flask import Flask, jsonify, request
import pandas as pd

app = Flask(__name__)
df = pd.read_excel('Bank_Employees.xlsx')

@app.route('/')
def index():
    return 'Welcome to the server'

@app.route('/bank', methods=['GET'])
def get_bank_data():
    global df
    data = df.to_dict(orient='records')
    return jsonify(data)

@app.route('/add_data', methods=['POST'])
def add_bank_data():
    global df

    # Get the JSON data from the POST request
    new_data = request.json

    # Convert the dictionary to a DataFrame
    new_data_df = pd.DataFrame([new_data])

    # Append the new data DataFrame to the existing DataFrame
    df = pd.concat([df, new_data_df], ignore_index=True)

    # Save the updated DataFrame to the Excel file
    df.to_excel('C:\\Users\\Admin\\PycharmProjects\\mycv\\Bank_Employees.xlsx', index=False)

    return jsonify({"message": "Data added successfully"})

@app.route('/update_data/<int:index>', methods=['PUT'])
def update_bank_data(index):
    global df

    # Get the JSON data from the PUT request
    updated_data = request.json

    # Update the corresponding row in the DataFrame
    df.loc[index] = updated_data

    # Save the updated DataFrame to the Excel file
    df.to_excel('C:\\Users\\Admin\\PycharmProjects\\mycv\\Bank_Employees.xlsx', index=False)

    return jsonify({"message": "Data updated successfully"})

@app.route('/delete_data/<int:index>', methods=['DELETE'])
def delete_bank_data(index):
    global df

    # Drop the specified row from the DataFrame
    df = df.drop(index)

    # Save the updated DataFrame to the Excel file
    df.to_excel('C:\\Users\\Admin\\PycharmProjects\\mycv\\Bank_Employees.xlsx',index=False)

    return jsonify({"message": "Data deleted successfully"})

if __name__ == '__main__':
    app.run(debug=True)
