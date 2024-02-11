# Exam Paper Generator

## Overview

This Python script generates randomized exam papers based on a set of questions stored in a Microsoft Word document. The script reads the questions, organizes them by type, and then selects a specified number of questions for each type to create an exam paper. The questions and answers are then written to text files, along with a timestamp.

## Features

- **Random Question Selection:** The script uses a random selection process to choose questions from different types, considering specified constraints for the maximum number of questions per type.
- **Function Decomposition:** The code is organized into functions for better readability and maintainability.
- **Timestamp Generation:** The exam papers are timestamped with the current date and time.
- **File Handling:** The script creates and writes the exam paper and answers to text files.

## Usage

1. **Install Dependencies:**
   - Make sure you have Python installed on your system.
   - Install required dependencies using the following command:
     ```bash
     pip install python-docx
     ```

2. **Run the Script:**
   - Execute the script by running the following command:
     ```bash
     python exam_paper_generator.py
     ```

3. **Output:**
   - The generated exam paper and answers will be stored in text files ('考试.txt' and '答案.txt', respectively).

## Configuration

- **Word Document:**
  - Update the `file_name` variable in the script with the name of your Microsoft Word document containing the questions.

## Notes

- This script assumes a specific structure in the Word document for organizing questions. Ensure that your document adheres to the expected format.

## Contributors

- Yingying Yu (Author)

## License

This project is licensed under the [MIT License](LICENSE).
