# Students Interdisciplinary Groups

This project is designed to assign students to interdisciplinary groups based on their course of study. The script reads student data from a CSV file, processes it, and outputs the group assignments to both CSV and Excel files.

## Requirements

- Python 3.x
- pandas
- openpyxl

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/RondinaMR/students-interdisciplinary-groups.git
    cd students-interdisciplinary-groups
    ```

2. Install the required Python packages:
   ```sh
    conda env create -f environment.yml
    ```
   or, using pip:
   ```sh
    pip install pandas openpyxl
    ```
    

## Usage

1. Place your `students.csv` file in the project directory. The CSV file should have the following columns:
    - `MATRICOLA`
    - `COGNOME - (*) Inserito dal docente`
    - `NOME`
    - `CDS STUDENTE`

2. Run the script:
    ```sh
    python main.py
    ```

3. The script will generate three output files:
    - `students-output.csv`
    - `students-output.xlsx`
    - `students-output-extended.xlsx`

## Explanation

- The script reads the student data from `students.csv`.
- It renames the column `COGNOME - (*) Inserito dal docente` to `COGNOME`.
- It assigns students to groups based on the number of students and the desired group size.
- It ensures diversity by grouping students from different courses.
- The output is saved in both CSV and Excel formats.
- The Excel files are formatted with alternating row colors for better readability.

## License

This project is licensed under the GNU General Public License v3.0.
