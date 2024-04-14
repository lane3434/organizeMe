"""
Script to organize tasks into an Excel file.
"""

def create_organize_excel(tasks):
    """
    Creates an Excel file with the provided tasks.

    Args:
        tasks (list): List of dictionaries containing task details.
            Each dictionary should have keys 'name', 'priority', and 'time'.
    """
    # Writing Excel file manually without using openpyxl
    with open("organize.xlsx", "w", encoding="utf-8") as file:
        file.write("Task,Priority,Time\n")

        for task in tasks:
            file.write(f"{task['name']},{task['priority']},{task['time']}\n")

    print("Excel file 'organize.xlsx' created successfully!")

def main():
    """
    Main function to prompt user for task details and create Excel file.
    """
    tasks = []
    while True:
        task_name = input("Enter task name (or type 'done' to finish): ")
        if task_name.lower() == 'done':
            break
        priority = input("Enter priority (high, medium, low): ")
        time = input("Enter estimated time (in minutes): ")

        tasks.append({'name': task_name, 'priority': priority, 'time': time})

    create_organize_excel(tasks)

if __name__ == "__main__":
    main()
