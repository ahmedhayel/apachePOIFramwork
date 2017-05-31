firstly you must download apache poi (bin)

example of using this framework is :

# `create file not exite with data :`

ApachePlugin poi = new ApachePlugin(String path, String fileName, String sheetName, List<anyThing> content, ColorsChoice colorHeaders,ColorsChoice colorContent)

## with `ColorsChoice is enumeration` 

poi.createNewExcelFile();  generate excel in the path

# `Editing existing file (empty) :`

poi.EditExistingExcelFile();


# `get data from existing file (not empty) :`

List<Object[]> objects=poi.getDataFromExcelFile(int numbreOfColumnsInTheFile);  the first array ==> header of table in Excel

