# Student Management System

A Python-based student information management system built with PyQt6. This project provides a user-friendly interface for managing student data, including user authentication, data visualization, and statistical analysis.

## Features
- User authentication with login attempts tracking
- Manage student information (add, edit, delete, search)
- Data visualization with pie charts and bar graphs
- Filter and sort data with advanced table headers
- Import and export student data in Excel format
- Modern, intuitive UI with icon buttons

## Installation

### Prerequisites
- Python >= 3.9
- pip (Python package manager)

### Clone the repository
```bash
git clone https://github.com/rachluis/Student-Management-System.git
cd Student-Management-System
```

### Install dependencies
```bash
pip install -r requirements.txt
```

## Usage

1. Prepare the user account file at `data/user.txt` in the format:
   ```
   username:password
   admin:admin123
   user1:password1
   ```
2. Prepare the student data Excel file at `data/学生数据示例.xlsx` (or rename as needed). The columns should be:
   - Name, Gender, Ethnicity, Department, Major, Province
3. Run the application:
   ```bash
   python main.py
   ```
4. Log in with your username and password.
5. Use the main window to add, edit, delete, search, and visualize student records.

## Dependencies
- pandas
- numpy
- PyQt6
- matplotlib
- pyqtgraph
- openpyxl (for Excel file support)

## Contribution
Contributions are welcome! Please open issues or submit pull requests for bug fixes, improvements, or new features.

## License
This project is licensed under the MIT License.
