# Excel导入工具

使用Wails和Go开发的Excel文件导入Oracle/MySQL数据库的桌面应用。

## 功能特性

- 支持Excel (.xlsx, .xls) 和 CSV 文件导入
- 支持MySQL和Oracle数据库
- 图形化用户界面
- 实时导入进度显示
- 错误处理和日志记录

## 项目结构

```
csv2o/
├── main.go           # Go后端主程序
├── go.mod           # Go模块文件
├── wails.json       # Wails配置文件
├── frontend/        # 前端代码
│   ├── src/
│   │   ├── main.tsx
│   │   ├── App.tsx
│   │   └── index.html
│   ├── package.json
│   ├── tsconfig.json
│   └── webpack.config.js
└── build/           # 构建输出目录
```

## 安装依赖

### 1. 安装Go依赖

```bash
go mod tidy
go get github.com/go-sql-driver/mysql
go get github.com/sijms/go-ora/v2
go get github.com/xuri/excelize/v2
go get github.com/wailsapp/wails/v2
```

### 2. 安装前端依赖

```bash
cd frontend
npm install
```

## 构建和运行

### 开发模式

```bash
# 启动前端开发服务器
cd frontend
npm run dev

# 在另一个终端启动Go后端
wails dev
```

### 生产构建

```bash
# 构建前端
cd frontend
npm run build

# 构建完整应用
wails build
```

## 使用方法

1. 启动应用
2. 选择数据库类型 (MySQL/Oracle)
3. 填写数据库连接信息
4. 选择要导入的Excel/CSV文件
5. 指定目标表名
6. 点击"开始导入"

## CSV文件格式

CSV文件第一行为标题行，格式如下：

```csv
ID,NAME,EMAIL,PHONE,ADDRESS
1,张三,zhangsan@example.com,13800138001,北京市朝阳区
2,李四,lisi@example.com,13800138002,上海市浦东新区
```

## 数据库要求

### MySQL
- 确保表已存在
- 字段名应与CSV标题匹配（不区分大小写）

### Oracle
- 确保表已存在
- 字段名应与CSV标题匹配（不区分大小写）

## 配置说明

### wails.json
```json
{
  "name": "Excel导入工具",
  "outputfilename": "excel-importer",
  "frontend:install": "npm install",
  "frontend:build": "npm run build",
  "frontend:dev:watcher": "npm run dev",
  "author": {
    "name": "Your Name",
    "email": "your@email.com"
  }
}
```

## 技术栈

- **后端**: Go
- **前端**: React + TypeScript
- **框架**: Wails v2
- **数据库**: MySQL, Oracle
- **文件处理**: Excelize (Excel), encoding/csv (CSV)

## 注意事项

1. 确保数据库表已存在
2. CSV文件编码应为UTF-8
3. Excel文件应为.xlsx或.xls格式
4. 确保数据库用户有插入权限

## 故障排除

### 常见问题

1. **数据库连接失败**
   - 检查数据库服务器是否运行
   - 验证连接参数（主机、端口、用户名、密码）

2. **文件读取失败**
   - 确保文件存在且有读取权限
   - 检查文件格式是否正确

3. **导入失败**
   - 检查表结构是否匹配
   - 验证数据类型是否兼容

## 开发说明

当前版本是一个演示版本，包含以下限制：
- 使用标准库处理CSV文件
- 模拟数据库操作
- 前端界面已实现但需要Wails框架支持

要获得完整功能，请安装所有依赖包并配置真实的数据库连接。