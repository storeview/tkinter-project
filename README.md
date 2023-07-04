# tkinter-project

## excel_file_merge

主要功能：

接收多个xls、xlsx文件，然后选择特定的几列，将几份文件中的数据保存到一个Excel表格中，并且具有筛选功能。

作用：

从云平台上，导出不同模块的Bug缺陷表格，通常是4-5个文件。对这些文件进行整理，数据汇总、筛选和总结是一件比较耗时耗力的工作。

例如，安全周报的数据需要从多个小组的Bug空间导出数据然后进行统计，还需要删除不需要的字段，保留特定的字段，以及筛选特定范围的数据行，**手工操作**非常繁琐。有了我这个工具，就可以从多个小组的Bug空间中导出数据后，再导入到工具里，一次性操作所有文件，并统一输出到一份Excel表格当中。方便进行记录及整理

### todo

- OK 将【源文件】的行宽和样式保存下来
- OK 在【输出文件】中自动将空白列删除，被选中的列依次排列，不能有空的一列
- 添加xls文件的处理模块

## cloudtest_testcase_sync

主要功能：

cloudtest平台上的测试用例，和MindMaster上的测试用例进行同步

操作流程：

1. cloudtest导出Excel格式，MindMaster导出Excel格式（另一种排布方式）
2. 将两份文件，分别选择到工具中
    - cloudtest_excel -> freemind(.mm)格式；mindmaster_excel -> freemind(.mm)格式
3. 工具开始比较，输出两份文件与【归档文件】的差别。在列表中显示cloudtest修改的内容和mindmaster中修改的内容
4. 如果点击同步cloudtest，则更改【归档文件】后，输出一份新的【归档文件】和mindmaster导入文件；如果点击同步mindmaster，则更改【归档文件】后，输出一份新的【归档文件】和cloudtest导入文件。如果双方都点击了同步，则输出一份新的【归档文件】、mindmaster导入文件、cloudtest导入文件。
5. 将输出的文件导入到各自软件当中，完成同步