example of using this framework is :

# `create file not exite with data :`

ApachePlugin poi = new ApachePlugin(String path, String fileName, String sheetName,List<String> headers, List<anyThing> content, String title)

poi.createNewExcelFile();  generate excel in the path

# `Editing existing file (must be empty) :`

poi.EditExistingExcelFile();


# `get data from existing file (not empty) :`

List<Object[]> objects=poi.getDataFromExcelFile(int numbreOfColumnsInTheFile);  the first array ==> header of table in Excel

List<Object[]> ob = poi.getDataFromExcelFile(5); for example the file has 5 columns

		for (Object[] p : ob) {
			for (int i = 0; i < p.length; i++) {

				System.out.print(p[i] + " \t\t ");
			}
			System.out.println();
		}

