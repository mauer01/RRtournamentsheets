# Round Robin Tournament Calc Sheets

This AutoIt script generates OpenOffice Calc sheets for 2-game round-robin tournaments based on a formatted text file input.

## Features

- Automatically generates OpenOffice Calc (ODS) files for each tournament.
- Supports multiple tournaments in a single input file.
- Generates 2-game round-robin schedules for each tournament.
- Populates the Calc sheets with player names and results format.

## Requirements

- [AutoIt 3.3.16.1](https://www.autoitscript.com/site/autoit/downloads/) -only if you wanna compile it yourself.
- OpenOffice or LibreOffice (for Calc)

## Usage

1. **Input File Format:**
    - The input text file should be formatted as follows:
    ```
    :*tournament1*
    *player1*
    *player2*
    *player3*
    ...
    :*tournament2*
    *player1*
    *player2*
    *player3*
    ...
    ```

2. **Running the Script:**
    - You can run the script by either:
        - Dragging and dropping the text file onto the executable.
        - Running from the command line:
        ```
        round robin open calc sheets.exe path_to_textfile.txt
        ```
        - If no file is specified, a file dialog will prompt you to select one.

3. **Output:**
    - The script will generate an ODS file in the script directory with the name format `HHMMSS_tournament.ods`.

## Functions

### Main Functions

- **constructsheet($HNDL_ODS, $tournament_name, $players, $playernames)**
    - Creates a new sheet for a tournament.
    - Writes player names and sets up the results format.
  
- **GenerateRoundRobinSchedule($teams)**
    - Generates the round-robin schedule.
    - Handles odd numbers of teams by adding a dummy team.

### Helper Functions

- **_OOoCalc_WriteRowFromArray($row, $x, $y)**
    - Writes a row of data starting from cell ($x, $y).
  
- **_OOoCalc_DraggingDown($range, $count)**
    - Copies a cell range down by a specified count.
  
- **ShuffleArray(ByRef $array)**
    - Shuffles an array in place.
  
- **AlphabeticalNumberToLetter($number)**
    - Converts a number to its corresponding alphabetical letter.
  
- **LetterToAlphabeticalNumber($letter)**
    - Converts a letter to its corresponding alphabetical number.

- **Twentysixintodec($number)**
    - Converts a base-26 string to a decimal number.

## Error Handling

- Ensures the input file is correctly formatted before processing.
- Displays an error message if the input file is not set up correctly.

## License

This project is licensed under the MIT License.

## Author

- Mauer01

For any issues or questions, please open an issue on this GitHub repository.
