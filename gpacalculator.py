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

def read_course_data(course):
    course.name = input("Enter course name: ")
    course.score = float(input("Enter course score: "))
    course.credit_hours = int(input("Enter credit hours: "))

def calculate_gpa(total_grade_points, total_credit_hours):
    return total_grade_points / total_credit_hours if total_credit_hours != 0 else 0

def display_course_data(course):
    print(f"{course.name}: {course.score}, {course.credit_hours}")

def find_course_by_name(course_list, name):
    for course in course_list:
        if course.name == name:
            return course
    return None

def add_course(course_list):
    new_course = Course()
    read_course_data(new_course)

    if find_course_by_name(course_list, new_course.name):
        print(f"Course with name '{new_course.name}' already exists. Skipping addition.")
    else:
        course_list.append(new_course)
        print(f"Course '{new_course.name}' added.")

def remove_course(course_list, name):
    course = find_course_by_name(course_list, name)
    if course:
        course_list.remove(course)
        print(f"Course '{name}' removed.")
    else:
        print(f"Course '{name}' not found.")

def access_course(course_list):
    course_name = input("Enter the name of the course to access: ")
    course = find_course_by_name(course_list, course_name)
    if course:
        print("\nOptions:")
        print("1) Display course data")
        print("2) Edit course data")
        option = input("Enter option (1 or 2): ")

        if option == "1":
            display_course_data(course)
        elif option == "2":
            read_course_data(course)
            print("Modified course details:")
            display_course_data(course)
        else:
            print("Invalid option.")
    else:
        print(f"Course '{course_name}' not found.")

def list_all_courses(course_list):
    print("\nAll Courses:")
    for course in course_list:
        display_course_data(course)

def save_to_file(course_list):
    with open("user_data.txt", "w") as file:
        for course in course_list:
            file.write(f"{course.name}\n{course.score}\n{course.credit_hours}\n")

def load_from_file(course_list):
    if os.path.exists("user_data.txt"):
        with open("user_data.txt", "r") as file:
            while True:
                name = file.readline().strip()
                if not name:
                    break
                score = float(file.readline().strip())
                credit_hours = int(file.readline().strip())
                course_list.append(Course(name, score, credit_hours))

def main():
    course_list = []
    load_from_file(course_list)

    while True:
        print("\nOptions:")
        print("1) Calculate GPA")
        print("2) Add course")
        print("3) Remove course")
        print("4) Access course")
        print("5) List all courses")
        print("6) EndProgram")
        option = input("Enter option (1-6): ")

        if option == "1":
            total_credit_hours = calculate_total_credit_hours(course_list)
            total_grade_points = calculate_total_grade_points(course_list)
            gpa = calculate_gpa(total_grade_points, total_credit_hours)
            print(f"Your GPA is: {gpa}")
        elif option == "2":
            add_course(course_list)
        elif option == "3":
            course_name = input("Enter the name of the course to remove: ")
            remove_course(course_list, course_name)
        elif option == "4":
            access_course(course_list)
        elif option == "5":
            list_all_courses(course_list)
        elif option == "6":
            save_to_file(course_list)
            print("Program ended.")
            break
        else:
            print("Invalid option. Please enter a number between 1 and 6.")

if __name__ == "__main__":
    main()
