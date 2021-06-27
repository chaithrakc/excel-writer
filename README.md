# EXCEL WRITER— JAVA LIBRARY

We often require the data in the excel files (preferred format), which is used by the Managers, Business Analysts, and others for viewing and manipulating the data to gain more insight. 
Excel files make the data helpful for effective decision making. This Excel Writer Java library eliminates boilerplate code, written in automation tools that generate reports. It reduces approximately 100-150 lines of boilerplate code. Developers can focus on the implementation of business logic rather than coding for report generation.

# How to use?
Add this Java library in the automations, which has the requirement for report generation.

Sample Code:
```
List<CategoryVO> categories = categorization.getCategories();
ExcelWriter<CategoryVO> excelWriter = new ExcelWriter<>("C:\\CategoryReport.xlsx", CategoryVO.class);
excelWriter.renderData(categories);
excelWriter.generateExcelFile(); 
```
`categories` will be written into the `CategoryReport.xlsx` file.

Sample Report (`CategoryReport.xlsx`):
```
| Provider    | Firm | Branch | ErrorCode     | ErrorDescription | Description   |
|-------------|------|--------|---------------|------------------|---------------|
| xxxx-yy-prd | zzzz | Main   | 0             | Not an Error     | OK            |
| xxxx-yy-prd | zzzz | Main   | 0             | Not an Error     | OK            |
| xxxx-yy-prd | zzzz | Main   | 0             | Not an Error     | OK            |
| xxxx-yy-prd | zzzz | Main   | 0             | Not an Error     | OK            |
| xxxx-yy-prd | zzzz | Main   | 0             | Not an Error     | OK            |
| xxxx-yy-prd | zzzz | Main   | 0             | Not an Error     | OK            |
| xxxx-yy-prd | zzzz | Main   | 0             | Not an Error     | OK            |
| xxxx-yy-prd | zzzz | Main   | Not Available | Not Available    | Not Available |
```

# Future Scope
<ol type="a">
<li> Pagination in case of application wants to write data in parts</li>
<li> Writing the data present in nested classes and data structures (such as Map, List, and Set) into the report file</li>
<li> Generating multiple excel files if data exceeds the maximum limit of excel file size</li>
</ol>

