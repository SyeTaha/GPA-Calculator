import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import pandas as pd
import os

class Course:
    def __init__(self, name="", score=0, credit_hours=0):
        self.name = name
        self.score = score
        self.credit_hours = credit_hours

def calculate_total_credit_hours(course_list):
    return sum(course.credit_hours for course in course_list)

def calculate_total_grade_points(course_list):
    return sum(course.credit_hours * course.score for course in course_list)

def calculate_gpa(total_grade_points, total_credit_hours):
    return total_grade_points / total_credit_hours if total_credit_hours != 0 else 0

def add_course(course_list, tree, current_gpa_label):
    top = tk.Toplevel()
    top.title("Add Course")
    top.geometry("300x200")  

    tk.Label(top, text="Course Name:").pack()
    entry_name = tk.Entry(top)
    entry_name.pack()

    tk.Label(top, text="Score:").pack()
    entry_score = tk.Entry(top)
    entry_score.pack()

    tk.Label(top, text="Credit Hours:").pack()
    entry_credit_hours = tk.Entry(top)
    entry_credit_hours.pack()

    add_button = tk.Button(top, text="Add Course", command=lambda: process_add_course(course_list, tree, current_gpa_label, entry_name, entry_score, entry_credit_hours, top))
    add_button.pack()

def process_add_course(course_list, tree, current_gpa_label, entry_name, entry_score, entry_credit_hours, top):
    new_course = Course()
    new_course.name = entry_name.get()
    
    try:
        new_course.score = float(entry_score.get())
        new_course.credit_hours = int(entry_credit_hours.get())
    except ValueError:
        messagebox.showerror("Error", "Invalid input. Please enter valid numbers.")
        return


    if find_course_by_name(course_list, new_course.name):
        messagebox.showwarning("Warning", f"Course with name '{new_course.name}' already exists. Skipping addition.")
    else:
        course_list.append(new_course)
        tree.insert("", tk.END, values=(new_course.name, new_course.score, new_course.credit_hours))
        update_gpa_and_show_message(course_list, current_gpa_label, f"Course '{new_course.name}' added.")
        top.destroy()

def remove_course(course_list, tree, current_gpa_label):
    selected_item = tree.selection()
    if selected_item:
        course_data = tree.item(selected_item, 'values')
        course_name = course_data[0]
        course = find_course_by_name(course_list, course_name)
        if course:
            course_list.remove(course)
            tree.delete(selected_item)
            update_gpa_and_show_message(course_list, current_gpa_label, f"Course '{course_name}' removed.")
        else:
            messagebox.showwarning("Warning", f"Course '{course_name}' not found.")
    else:
        messagebox.showwarning("Warning", "Please select a course to remove.")

def edit_course(course_list, tree, current_gpa_label):
    selected_item = tree.selection()
    if selected_item:
        course_data = tree.item(selected_item, 'values')
        course_name = course_data[0]
        course = find_course_by_name(course_list, course_name)
        if course:
            edit_course_window(course, tree, selected_item, current_gpa_label)
        else:
            messagebox.showwarning("Warning", f"Course '{course_name}' not found.")
    else:
        messagebox.showwarning("Warning", "Please select a course to edit.")

def edit_course_window(course, tree, index, current_gpa_label):
    top = tk.Toplevel()
    top.title("Edit Course")
    top.geometry("300x200")  

    tk.Label(top, text="Course Name:").pack()
    entry_name = tk.Entry(top)
    entry_name.pack()
    entry_name.insert(0, course.name)

    tk.Label(top, text="Score:").pack()
    entry_score = tk.Entry(top)
    entry_score.pack()
    entry_score.insert(0, str(course.score))

    tk.Label(top, text="Credit Hours:").pack()
    entry_credit_hours = tk.Entry(top)
    entry_credit_hours.pack()
    entry_credit_hours.insert(0, str(course.credit_hours))

    tk.Button(top, text="Update", command=lambda: update_course(course, entry_name, entry_score, entry_credit_hours, top, tree, index, current_gpa_label)).pack()

def update_course(course, entry_name, entry_score, entry_credit_hours, top, tree, index, current_gpa_label):
    course.name = entry_name.get()
    try:
        course.score = float(entry_score.get())
        course.credit_hours = int(entry_credit_hours.get())
        tree.item(index, values=(course.name, course.score, course.credit_hours))
        top.destroy()
    except ValueError:
        messagebox.showerror("Error", "Invalid input. Please enter valid numbers.")

def save_to_excel(course_list):
    folder_path = 'C:\\CourseData'

    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    file_path = os.path.join(folder_path, 'user_data.xlsx')

    data = {'Name': [course.name for course in course_list],
            'Score': [course.score for course in course_list],
            'Credit Hours': [course.credit_hours for course in course_list]}

    df = pd.DataFrame(data)
    df.to_excel(file_path, index=False)
    update_gpa_and_show_message(course_list, None, "Data saved to 'user_data.xlsx'.")

def load_from_excel(course_list):
    folder_path = 'C:\\CourseData'
    file_path = os.path.join(folder_path, 'user_data.xlsx')

    if os.path.exists(file_path):
        df = pd.read_excel(file_path)
        for index, row in df.iterrows():
            course_list.append(Course(row['Name'], row['Score'], row['Credit Hours']))

    else:
        update_gpa_and_show_message(course_list, None, f"File '{file_path}' not found. Creating a new one.")

def find_course_by_name(course_list, name):
    for course in course_list:
        if course.name == name:
            return course
    return None

def display_course_data(course):
    messagebox.showinfo("Course Data", f"{course.name}: {course.score}, {course.credit_hours}")

def list_all_courses(course_list, tree, current_gpa_label):
    tree.delete(*tree.get_children())
    for course in course_list:
        tree.insert("", tk.END, values=(course.name, course.score, course.credit_hours))

    update_gpa_and_show_message(course_list, current_gpa_label)

def update_gpa_and_show_message(course_list, current_gpa_label, message=None):
    if current_gpa_label:
        total_credit_hours = calculate_total_credit_hours(course_list)
        total_grade_points = calculate_total_grade_points(course_list)
        gpa = calculate_gpa(total_grade_points, total_credit_hours)
        current_gpa_label.config(text=f"Current GPA: {gpa:.2f}")

    if message:
        messagebox.showinfo("Information", message)

def main():
    folder_path = 'C:\\CourseData'

    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    course_list = []
    load_from_excel(course_list)

    def calculate_gpa_command():
        update_gpa_and_show_message(course_list, current_gpa_label)

    root = tk.Tk()
    root.title("GPA Calculator")
    root.geometry("350x350")

    gpa_frame = tk.Frame(root)
    gpa_frame.pack(pady=10)

    calculate_button = tk.Button(gpa_frame, text="Calculate GPA", command=calculate_gpa_command, relief=tk.GROOVE, bg="#4CAF50", fg="white", font=('Helvetica', 10, 'bold'))
    calculate_button.grid(row=0, column=0, padx=5)

    current_gpa_label = tk.Label(gpa_frame, text="Current GPA: ")
    current_gpa_label.grid(row=0, column=1, padx=5)

    save_button = tk.Button(gpa_frame, text="Save", command=lambda: save_to_excel(course_list), relief=tk.GROOVE, bg="#2196F3", fg="white", font=('Helvetica', 10, 'bold'))
    save_button.grid(row=0, column=2, padx=5)

    courses_frame = tk.Frame(root)
    courses_frame.pack()

    columns = ("Course Name", "Score", "Credit Hours")
    courses_tree = ttk.Treeview(courses_frame, columns=columns, show="headings", height=10)
    courses_tree.pack(side=tk.LEFT, padx=10)

    for col in columns:
        courses_tree.heading(col, text=col)
        courses_tree.column(col, width=100)

    add_button = tk.Button(root, text="+", command=lambda: add_course(course_list, courses_tree, current_gpa_label), width=5, relief=tk.GROOVE, bg="#4CAF50", fg="white", font=('Helvetica', 10, 'bold'))
    add_button.pack(side=tk.RIGHT, padx=5)

    remove_button = tk.Button(root, text="-", command=lambda: remove_course(course_list, courses_tree, current_gpa_label), width=5, relief=tk.GROOVE, bg="#F44336", fg="white", font=('Helvetica', 10, 'bold'))
    remove_button.pack(side=tk.RIGHT, padx=5)

    edit_button = tk.Button(root, text="Edit", command=lambda: edit_course(course_list, courses_tree, current_gpa_label), relief=tk.GROOVE, bg="#2196F3", fg="white", font=('Helvetica', 10, 'bold'))
    edit_button.pack(side=tk.RIGHT, padx=5)

    list_button = tk.Button(root, text="List All Courses", command=lambda: list_all_courses(course_list, courses_tree, current_gpa_label), relief=tk.GROOVE, bg="#607D8B", fg="white", font=('Helvetica', 10, 'bold'))
    list_button.pack(side=tk.LEFT, padx=5)

    root.mainloop()

if __name__ == "__main__":
    main()
