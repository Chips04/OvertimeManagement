# OvertimeManagement

**第一次使用**

右键-打开终端-输入下边命令行

sudo apt install python3-pip

pip install openpyxl pandas

sudo apt update

sudo apt install git

新建一个文件夹放代码，在这个文件夹中执行下列命令（用户名和邮箱地址写你自己的）：

git config --global user.name "用户名"

git config --global user.email "邮箱地址"

git clone https://github.com/Chips04/OvertimeManagement.git

1.加班表格处理文件夹可以粘出来放到你喜欢的地方，如桌面

2.variables.txt文件要放到OvertimeManagement/pythonProject文件夹下，其中的path是你的加班表格处理文件夹所在路径，路径最后要有一个斜杠（重要），里面还有要生成的表格的年月和需要的组别（生成调休表和年假表使用），需要手动进行修改（一定要改），文本编辑器修改即可。

**撤销全部本地修改**

git restore .
（如果你修改了本地文件，拉代码时会冲突，执行此命令后再拉就不会有问题了）

**拉最新代码**

git pull origin main

**忽略文件**

git rm --cached pythonProject/variables.txt

**处理表格**

①处理数据.py

平台上，选择加班申请表单，权限组选择“管理全部流程2”，选中“处理上个月全部的数据”视图，在右侧筛选处添加过滤条件：所属组别 等于 综合组。点击筛选。然后导出筛选后的数据。把这个导出的文件粘到“加班表格处理”文件夹下，删除此文件夹下原来的加班申请_balabala文件。

在全是py文件那里(OvertimeManagement/pythonProject/)文件夹下打开终端，敲下列命令执行：

python3 处理数据.py

会生成：加班费申报表、加班费发放表、加班费审批表、班日志汇总表、加班费申报表（第x季度）、领导画×表

②处理调休表.py

1.执行①后的操作以后，回到平台，点击右边筛选按钮，把加班日期 动态筛选 上月这条删掉。点击筛选。然后导出筛选后的数据。把这个导出的文件粘到“加班表格处理/调休处理数据源”文件夹下，删除此文件夹下原来的加班申请_balabala文件。

2.平台上，选择请休假-上芬应用，点击人员管理-请休假表单，导出全部数据。把这个导出的文件粘到“加班表格处理”文件夹下，删除此文件夹下原来的人员管理-请休假_balabala文件。

在全是py文件那里打开终端，敲下列命令执行：

python3 处理调休表.py

会生成：补休情况登记表-全、补休情况登记表-几月、补休情况登记表-几月-大表

③处理年假表.py

在全是py文件那里打开终端，敲下列命令执行：

python3 处理年假表.py

会生成：年假处理表

