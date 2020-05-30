## Test best library and solution to use when exporting very large data to excel files in .NET Framework
## Rerferences
### project & test documentation
+ [.Net 大数据量导出Excel方案](https://www.jianshu.com/p/23250c9d3684)

### knowledge references
+ [Writing Large Excel Files with the Open XML SDK](https://docs.microsoft.com/en-us/archive/blogs/brian_jones/writing-large-excel-files-with-the-open-xml-sdk)
+ [Export big amount of data from XLSX - OutOfMemoryException](https://stackoverflow.com/questions/32690851/export-big-amount-of-data-from-xlsx-outofmemoryexception)
+ [How to properly use OpenXmlWriter to write large Excel files](http://polymathprogrammer.com/2012/08/06/how-to-properly-use-openxmlwriter-to-write-large-excel-files/)
+ [Working with the shared string table (Open XML SDK)](https://docs.microsoft.com/en-us/office/open-xml/working-with-the-shared-string-table)

## Conclusion
### NPOI、ClosedXML etc
these libraries will keep excel data in memory, so if the DataTable to export has very large scales(Let's say more than 600,000+), the memory usage will be extremely high
### OpenXML
this library actually has the best performance, especially when using OpenXMLWriter directly