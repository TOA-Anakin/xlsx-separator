# xlsx-separator

Generating separate Excel files (`.xlsx`) files from the input Excel file based on values of one of its columns.

1. Install composer packages: `composer install`

2. Place the input Escel file into the `/input` folder.

3. Run the local PHP server on port 8888: `php -S localhost:8888 -t src`

4. In order to generate separate Excel files, visit `http://localhost:8888` in the browser. New Excel files will then be generated in the `/output` folder.