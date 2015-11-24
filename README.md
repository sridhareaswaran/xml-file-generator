### Multiple xml file generator 
Excel VBA utility used to generate multiple xml files from the excel data.

> Handly when you need to create multiple files for **data-driven tests** & stuff.

#### Screenshot: 
![alt text][logo]

#### Usage:
- **data** sheet
	- First row is your xml element name.
	- Place your data for respective element in the particular column.
	- If your element has attributes value, do give it in the first row itself. (working on making attributes value applicable at file level)
- **Dashboard** sheet
	- Enter in the file type. (Even though it generates only xml data, provided a way to save those data in xml/text format)
	- Enter the file location to be saved.
	- Enter the Start & End node for xml. (Not mandatory, but gives you proper xml syntax.)

#### To-dos:
- Make attributes value for elements applicable at file level.

> cheers :)

[logo]: https://raw.githubusercontent.com/sridhareaswaran/xml-file-generator/master/img/xml%20generator.png "xml generator"