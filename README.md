# PowerQuery_relative-path-setting
PowerQuery by default queries via absolute path, which can be problematic when running the query on a different machine. This a my note of how to set up relative path query in MS Excel PowerQuery. So that regardless of where the query file is saved, as long as the relative location of data source file stays the same, the query should operate as normal.

Below is an example of querying one single data source file in the same folder. For scenarios that the source file may be in a subfolder, or the source IS every file in the subfolder, the codes below will require some tweeking. I will find time to add examples for those, but for now:

find any cell and use below formula in Excel to return current the path

    =LEFT(CELL("filename",$A$1),FIND("[",CELL("filename",$A$1),1)-1)

![image](https://user-images.githubusercontent.com/117622597/231323028-c147bb83-dc07-42de-86d2-f936b4bc8947.png)

use name namager to name above cell as FILE_PATH

In query editor, open Advanced Editor, and replace the source code with:

    let

        FILE_PATH = Excel.CurrentWorkbook(){[Name="FILE_PATH"]}[Content]{0}[Column1],
    
        FullPathToFile = FILE_PATH & "(filename).xlsx",

        Source = Excel.Workbook(File.Contents(FullPathToFile), null, true),
