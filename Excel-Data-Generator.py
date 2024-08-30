import csv
import pandas as pd
import os

# Add data to CSV File
def add_data(filename):
    with open(filename, mode='a', newline='') as file:
        writer = csv.writer(file)
        name = input("Enter name: ")
        age = input("Enter age: ")
        enroll = input("Enter enrollment number: ")
        writer.writerow([name, age, enroll])
        print("Data added successfully!")

# Display data From CSV File
def display_data(filename):
    with open(filename, mode='r') as file:
        reader = csv.reader(file)
        for row in reader:
            print(row)

# Update Data
def update_data(filename):
    temp_file = 'temp.csv'
    with open(filename, mode='r') as file, open(temp_file, mode='w', newline='') as temp:
        reader = csv.reader(file)
        writer = csv.writer(temp)
        enroll_to_update = input("Enter enrollment number to update: ")
        updated = False
        for row in reader:
            if row[2] == enroll_to_update:
                name = input("Enter new name: ")
                age = input("Enter new age: ")
                writer.writerow([name, age, enroll_to_update])
                updated = True
            else:
                writer.writerow(row)
        if updated:
            print("Data updated successfully!")
        else:
            print("Enrollment number not found.")
    os.replace(temp_file, filename)

# Delete Data
def delete_data(filename):
    temp_file = 'temp.csv'
    with open(filename, mode='r') as file, open(temp_file, mode='w', newline='') as temp:
        reader = csv.reader(file)
        writer = csv.writer(temp)
        enroll_to_delete = input("Enter enrollment number to delete: ")
        deleted = False
        for row in reader:
            if row[2] != enroll_to_delete:
                writer.writerow(row)
            else:
                deleted = True
        if deleted:
            print("Data deleted successfully!")
        else:
            print("Enrollment number not found.")
    os.replace(temp_file, filename)

# Convert CSV to Excel
def convert_csv_to_excel(csv_filename, excel_filename):
    df = pd.read_csv(csv_filename)
    df.to_excel(excel_filename, index=False)
    print(f"CSV file converted to Excel file: {excel_filename}")

# Main function
def main():
    filename = 'data.csv'
    if not os.path.exists(filename):
        with open(filename, mode='w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(['Name', 'Age', 'Enrollment Number'])
    
    while True:
        print("--------------------------------")
        print("--------------------------------")
        print("\tOPERATIONS MENU")
        print("\t1. Add Data")
        print("\t2. Display Data")
        print("\t3. Update Data")
        print("\t4. Delete Data")
        print("\t5. Convert CSV to Excel")
        print("\t6. Exit")
        print("--------------------------------")
        print("--------------------------------")
        choice = input("Enter your choice: ")
        
        if choice == '1':
            add_data(filename)
        elif choice == '2':
            display_data(filename)
        elif choice == '3':
            update_data(filename)
        elif choice == '4':
            delete_data(filename)
        elif choice == '5':
            convert_csv_to_excel(filename, 'GenData.xlsx')
        elif choice == '6':
            break
        else:
            print("Invalid choice. Please try again.")

if __name__ == "__main__":
    main()
