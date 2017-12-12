1. 首先文档要保存成.xls格式，不是.xlsx
2. 所有错题序号的填写除了数字、空格还有'-'符号外，其他字符一律不准出现， 针对这一条会后期改进
3. 所有合并的单元格的值的获取，为第一行/列的值
4. 设置配置文件， 对不同的需求设置对应的参数，进一步提高程序的易用性
5. getlastlineNums 函数要根据表格的具体格式改写
6. font/alignment 可以放在配置文件中
7. font/XFStyle/Alignment 的使用说明 -> 文档
8. xlsxwriter 同xlwt 只能创建新文件，不能修改原文件，但是可以处理xlsx
9. 程序中文件的读取路径使用os.path
10. 列出安装包，方便平台移植到windows上
11. .cfg文件中的注释 使用#或; 但是不可以和item 在一行
12. 周六上午的模板 Q/R两列是在一起的
13. 单元格自适应
14. virtualenv 方便移植


配置文件（.cfg）
# 参数:
#     ansLibName = ""
#     fileList = []             # 需要处理的文件的文件名的列表
#     startRow = int            # 学生信息的首行行号索引
#     startCol = int            # 保存课前测错题的（合并）单元格的（第一列）列号索引
#     ansCol = int              # 点拨放置的（合并）单元格的（第一列）列号索引
#     gradeCol = int            # 课前测成绩所在列的列号索引
#     rewardWords = ""          # 对于满分无错题的学生的表扬
#     finishTips = ""           # 程序完成课前测点拨内容填充后的提示！
