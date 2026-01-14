package main

import (
	"context"
	"database/sql"
	"embed"
	"encoding/csv"
	"encoding/json"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"
	"time"

	_ "github.com/go-sql-driver/mysql"
	_ "github.com/sijms/go-ora/v2"
	"github.com/wailsapp/wails/v2"
	"github.com/wailsapp/wails/v2/pkg/options"
	"github.com/wailsapp/wails/v2/pkg/options/assetserver"
	"github.com/wailsapp/wails/v2/pkg/runtime"
	"github.com/xuri/excelize/v2"
)

//go:embed all:frontend/dist
var assets embed.FS

// App struct
type App struct {
	ctx context.Context
}

// DBConfig 用于持久化数据库连接配置
type DBConfig struct {
	DbType         string `json:"dbType"`
	Host           string `json:"host"`
	Port           string `json:"port"`
	Database       string `json:"database"`
	Username       string `json:"username"`
	Password       string `json:"password"`
	TableName      string `json:"tableName"`
	ConnectionType string `json:"connectionType"`
	ServiceName    string `json:"serviceName"`
	TnsConnection  string `json:"tnsConnection"`
	TruncateChars  string `json:"truncateChars"`
}

type TableColumnInfo struct {
	ColumnName string
	DataType   string
	DataLength int
}

// NewApp creates a new App application struct
func NewApp() *App {
	return &App{}
}

// startup is called at application startup
func (a *App) startup(ctx context.Context) {
	// Perform your setup here
	a.ctx = ctx
	log.Println("Wails application started, backend is ready")
}

// domReady is called after the front-end dom has been loaded
func (a *App) domReady(ctx context.Context) {
	// Add your action here
	log.Println("DOM ready, frontend loaded")
}

// getConfigPath 返回配置文件路径
func getConfigPath() (string, error) {
	dir, err := os.UserConfigDir()
	if err != nil {
		return "", err
	}
	confDir := filepath.Join(dir, "csv2o")
	if err := os.MkdirAll(confDir, 0o755); err != nil {
		return "", err
	}
	return filepath.Join(confDir, "dbconfig.json"), nil
}

// SaveConfig 将配置持久化到本地
func (a *App) SaveConfig(cfg DBConfig) string {
	path, err := getConfigPath()
	if err != nil {
		return "错误: 无法获取配置路径: " + err.Error()
	}
	data, err := json.MarshalIndent(cfg, "", "  ")
	if err != nil {
		return "错误: 序列化配置失败: " + err.Error()
	}
	if err := os.WriteFile(path, data, 0o600); err != nil {
		return "错误: 写入配置文件失败: " + err.Error()
	}
	log.Println("配置已保存到", path)
	return "配置已保存"
}

// LoadConfig 读取本地配置
func (a *App) LoadConfig() DBConfig {
	var cfg DBConfig
	path, err := getConfigPath()
	if err != nil {
		log.Println("LoadConfig 获取路径失败:", err)
		return cfg
	}
	data, err := os.ReadFile(path)
	if err != nil {
		log.Println("LoadConfig 读取失败:", err)
		return cfg
	}
	if err := json.Unmarshal(data, &cfg); err != nil {
		log.Println("LoadConfig 解析失败:", err)
	}
	return cfg
}

// SelectExcelFile 使用原生文件对话框选择Excel/CSV文件，并返回完整路径
func (a *App) SelectExcelFile() string {
	if a.ctx == nil {
		log.Println("SelectExcelFile: context is nil")
		return ""
	}

	path, err := runtime.OpenFileDialog(a.ctx, runtime.OpenDialogOptions{
		Title: "选择Excel/CSV文件",
		Filters: []runtime.FileFilter{
			{
				DisplayName: "Excel文件 (*.xlsx, *.xls)",
				Pattern:     "*.xlsx;*.xls",
			},
			{
				DisplayName: "CSV文件 (*.csv)",
				Pattern:     "*.csv",
			},
		},
	})
	if err != nil {
		log.Printf("打开文件对话框失败: %v", err)
		return ""
	}
	log.Println("用户选择文件:", path)
	return path
}

// beforeClose is called when the application is about to quit,
// either by clicking the window close button or calling runtime.Quit.
// Returning true will cause the application to continue,
// false will continue shutdown as normal.
func (a *App) beforeClose(ctx context.Context) bool {
	return false
}

// shutdown is called at application termination
func (a *App) shutdown(ctx context.Context) {
	// Perform your teardown here
}

// Greet returns a greeting for the given name
func (a *App) Greet(name string) string {
	return fmt.Sprintf("Hello %s, It's show time!", name)
}

// GetExcelHeaders gets the header row from Excel/CSV file
func (a *App) GetExcelHeaders(filePath string) []string {
	// Read CSV file and return headers
	headers, err := readExcelHeaders(filePath)
	if err != nil {
		return []string{"错误: " + err.Error()}
	}
	return headers
}

// GetTableColumns gets table column information
func (a *App) GetTableColumns(dbType, host, port, username, password, tableName, connectionType, serviceName, tnsConnection string) []string {
	db, err := connectDatabase(dbType, host, port, username, password, connectionType, serviceName, tnsConnection)
	if err != nil {
		log.Printf("获取表结构时连接数据库失败: %v", err)
		return []string{"错误: " + err.Error()}
	}
	defer db.Close()

	tableName = strings.TrimSpace(tableName)
	if tableName == "" {
		return []string{"错误: 表名不能为空"}
	}

	var rows *sql.Rows

	switch strings.ToLower(dbType) {
	case "mysql":
		// 对于 MySQL，serviceName 作为数据库名使用
		schema := strings.TrimSpace(serviceName)
		if schema == "" {
			return []string{"错误: MySQL 需要提供数据库名"}
		}

		query := `
SELECT COLUMN_NAME
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ?
ORDER BY ORDINAL_POSITION`

		rows, err = db.Query(query, schema, tableName)
		if err != nil {
			log.Printf("查询 MySQL 表结构失败: %v", err)
			return []string{"错误: 查询表结构失败: " + err.Error()}
		}
	case "oracle":
		// 使用当前用户下的表
		query := `
SELECT COLUMN_NAME
FROM USER_TAB_COLUMNS
WHERE TABLE_NAME = UPPER(:1)
ORDER BY COLUMN_ID`

		rows, err = db.Query(query, tableName)
		if err != nil {
			log.Printf("查询 Oracle 表结构失败: %v", err)
			return []string{"错误: 查询表结构失败: " + err.Error()}
		}
	default:
		return []string{"错误: 不支持的数据库类型"}
	}
	defer rows.Close()

	var columns []string
	for rows.Next() {
		var col string
		if err := rows.Scan(&col); err != nil {
			log.Printf("扫描表结构字段失败: %v", err)
			return []string{"错误: 读取表结构失败: " + err.Error()}
		}
		columns = append(columns, strings.ToUpper(col))
	}

	if err := rows.Err(); err != nil {
		log.Printf("遍历表结构结果集失败: %v", err)
		return []string{"错误: 读取表结构失败: " + err.Error()}
	}
	fmt.Println(columns)
	return columns
}

// CompareFields compares Excel headers with database columns
func (a *App) CompareFields(excelHeaders []string, dbColumns []string) map[string]interface{} {
	result := make(map[string]interface{})

	// Convert to lowercase for case-insensitive comparison
	excelLower := make([]string, len(excelHeaders))
	dbLower := make([]string, len(dbColumns))

	excelMap := make(map[string]bool)
	dbMap := make(map[string]bool)

	for i, header := range excelHeaders {
		excelLower[i] = strings.ToLower(strings.TrimSpace(header))
		excelMap[excelLower[i]] = true
	}

	for i, col := range dbColumns {
		dbLower[i] = strings.ToLower(strings.TrimSpace(col))
		dbMap[dbLower[i]] = true
	}

	// Find matches, missing in DB, and extra in Excel
	var matched []string
	var missingInDb []string
	var extraInExcel []string

	for _, header := range excelLower {
		if dbMap[header] {
			matched = append(matched, header)
		} else {
			extraInExcel = append(extraInExcel, header)
		}
	}

	for _, col := range dbLower {
		if !excelMap[col] {
			missingInDb = append(missingInDb, col)
		}
	}

	result["matched"] = matched
	result["missingInDb"] = missingInDb
	result["extraInExcel"] = extraInExcel
	result["excelHeaders"] = excelHeaders
	result["dbColumns"] = dbColumns

	return result
}

func readExcelHeaders(filePath string) ([]string, error) {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("无法打开文件: %v", err)
	}
	defer f.Close()

	// 获取第一个工作表的名称
	sheetName := f.GetSheetName(0)

	// 使用流式迭代器读取行
	rows, err := f.Rows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("读取行失败: %v", err)
	}
	defer rows.Close()

	// 只迭代第一行
	if rows.Next() {
		columns, err := rows.Columns()
		if err != nil {
			return nil, err
		}
		return columns, nil
	}

	return nil, fmt.Errorf("文件内容为空")
}

// Helper function to read CSV headers
func readCsvHeaders(filePath string) ([]string, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return nil, fmt.Errorf("打开文件失败: %v", err)
	}
	defer file.Close()

	reader := csv.NewReader(file)
	headers, err := reader.Read()
	if err != nil {
		return nil, fmt.Errorf("读取文件头失败: %v", err)
	}

	return headers, nil
}

// ImportExcel imports data from Excel file to database
func (a *App) ImportExcel(dbType, host, port, username, password, tableName, filePath, connectionType, serviceName, tnsConnection, truncateChars string) string {
	batchSize := 1000
	enableTruncation := (truncateChars == "true")
	db, err := connectDatabase(dbType, host, port, username, password, connectionType, serviceName, tnsConnection)
	if err != nil {
		log.Printf("导入前连接数据库失败: %v", err)
		return "错误: 数据库连接失败: " + err.Error()
	}
	defer db.Close()

	var successCount int
	var totalExcelRows int

	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return fmt.Sprintf("无法打开Excel文件: %v", err.Error())
	}
	defer f.Close()

	// 获取工作表信息
	sheets := f.GetSheetMap()
	if len(sheets) == 0 {
		return "Excel文件不包含任何工作表"
	}

	sheetName := sheets[1] // 使用第一个工作表

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return fmt.Sprintf("读取工作表失败: %v", err.Error())
	}

	totalExcelRows = len(rows) - 1

	// 查询表结构
	query := `SELECT COLUMN_NAME, DATA_TYPE, DATA_LENGTH, NULLABLE
				  FROM ALL_TAB_COLUMNS
				  WHERE TABLE_NAME = UPPER(:1)
				  ORDER BY COLUMN_ID`

	res, err := db.Query(query, tableName)
	if err != nil {
		return fmt.Sprintf("查询表结构失败: %v", err.Error())
	}
	defer res.Close()

	var dbCols []TableColumnInfo
	for res.Next() {
		var c TableColumnInfo
		var nullable string
		if err := res.Scan(&c.ColumnName, &c.DataType, &c.DataLength, &nullable); err != nil {
			return fmt.Sprintf("解析列信息失败: %v", err.Error())
		}
		dbCols = append(dbCols, c)
	}

	if err := res.Err(); err != nil {
		return fmt.Sprintf("读取表结构时出错: %v", err.Error())
	}

	if len(dbCols) == 0 {
		return fmt.Sprintf("表 [%s] 不存在、无权限访问或不包含任何列", tableName)
	}

	// 字段匹配检查
	excelHeaders := rows[0]
	colMapping := make(map[string]int)
	var matchedCols, unmatchedCols []string

	for _, dbCol := range dbCols {
		found := false
		for idx, header := range excelHeaders {
			if strings.EqualFold(strings.TrimSpace(header), dbCol.ColumnName) {
				colMapping[dbCol.ColumnName] = idx
				matchedCols = append(matchedCols, dbCol.ColumnName)
				found = true
				break
			}
		}
		if !found {
			unmatchedCols = append(unmatchedCols, dbCol.ColumnName)
		}
	}

	if len(unmatchedCols) > 0 {
		return fmt.Sprintf("字段匹配失败: 缺少 %d 个必需字段", len(unmatchedCols))
	}

	// 准备 SQL 模板
	var placeholders []string
	for i, c := range dbCols {
		if strings.Contains(c.DataType, "DATE") {
			placeholders = append(placeholders, fmt.Sprintf("TO_DATE(:%d, 'YYYY-MM-DD HH24:MI:SS')", i+1))
		} else if enableTruncation && (strings.Contains(c.DataType, "VARCHAR") || strings.Contains(c.DataType, "CHAR")) && c.DataLength > 0 {
			// 使用Oracle的SUBSTRB函数进行字节级截断
			placeholders = append(placeholders, fmt.Sprintf("SUBSTRB(:%d, 1, %d)", i+1, c.DataLength))
		} else {
			placeholders = append(placeholders, fmt.Sprintf(":%d", i+1))
		}
	}
	insertSQL := fmt.Sprintf("INSERT INTO %s VALUES (%s)", tableName, strings.Join(placeholders, ","))

	dataRows := rows[1:]
	columnBuffers := make([][]interface{}, len(dbCols))
	for i := range columnBuffers {
		columnBuffers[i] = make([]interface{}, 0, batchSize)
	}

	// 批量刷新与错误探测逻辑
	flush := func(startIndex int) error {
		count := len(columnBuffers[0])
		if count == 0 {
			return nil
		}

		tx, _ := db.Begin()
		args := make([]interface{}, len(dbCols))
		for i := range columnBuffers {
			args[i] = columnBuffers[i]
		}

		_, err := tx.Exec(insertSQL, args...)
		if err != nil {
			tx.Rollback()

			// 找到第一个失败的行并立即返回
			for k := 0; k < count; k++ {
				singleArgs := make([]interface{}, len(dbCols))
				for cIdx := range dbCols {
					singleArgs[cIdx] = columnBuffers[cIdx][k]
				}

				if _, sErr := db.Exec(insertSQL, singleArgs...); sErr != nil {
					eLine := startIndex + k + 2
					// 直接记录错误信息，确保立即显示
					// 强制等待一小段时间确保日志显示
					time.Sleep(100 * time.Millisecond)
					return fmt.Errorf("数据库插入失败 (第%d行): %v", eLine, sErr)
				}
			}

			// 如果找不到具体错误行，返回原始错误
			return err
		}
		successCount += count
		return tx.Commit()
	}

	// 循环处理数据
	for i, row := range dataRows {
		for j, dbCol := range dbCols {
			idx := colMapping[dbCol.ColumnName]
			val := ""
			if idx < len(row) {
				val = strings.TrimSpace(row[idx])
			}

			// 处理不同数据类型的转换
			if strings.Contains(dbCol.DataType, "DATE") && val != "" {
				t, pErr := tryParseDate(val)
				if pErr != nil {
					return fmt.Sprintf("行 %d 日期格式不规范: %s", i+2, val)
				}
				columnBuffers[j] = append(columnBuffers[j], t.Format("2006-01-02 15:04:05"))
			} else if strings.Contains(dbCol.DataType, "NUMBER") && val == "" {
				columnBuffers[j] = append(columnBuffers[j], nil)
			} else {
				// 对于字符串类型，直接传递原始值，由Oracle的SUBSTRB函数处理截断
				columnBuffers[j] = append(columnBuffers[j], val)
			}
		}

		if (i+1)%batchSize == 0 || i == totalExcelRows-1 {
			sIdx := (i / batchSize) * batchSize
			if err := flush(sIdx); err != nil {
				return err.Error()
			}
			for j := range columnBuffers {
				columnBuffers[j] = columnBuffers[j][:0]
			}
		}
	}

	return fmt.Sprintf("excel行数:%d,成功导入:%d", totalExcelRows, len(dataRows))
}

// 智能日期转换
func tryParseDate(val string) (time.Time, error) {
	val = strings.TrimSpace(val)
	if val == "" {
		return time.Time{}, nil
	}
	formats := []string{
		"2006-01-02 15:04:05", "2006/1/2 15:04:05",
		"2006-01-02", "2006/1/2", "20060102", "02-Jan-06",
	}
	for _, f := range formats {
		if t, err := time.Parse(f, val); err == nil {
			return t, nil
		}
	}
	return time.Time{}, fmt.Errorf("无法识别日期格式: %s", val)
}

// TestDatabaseConnection tests database connectivity with Oracle/MySQL specific parameters
func (a *App) TestDatabaseConnection(dbType, host, port, username, password, connectionType, serviceName, tnsConnection string) string {
	db, err := connectDatabase(dbType, host, port, username, password, connectionType, serviceName, tnsConnection)
	if err != nil {
		log.Printf("数据库连接测试失败: %v", err)
		return "错误: 数据库连接失败: " + err.Error()
	}
	defer db.Close()

	desc := ""
	switch strings.ToLower(dbType) {
	case "mysql":
		desc = fmt.Sprintf("MySQL: %s:%s/%s", host, port, serviceName)
	case "oracle":
		switch connectionType {
		case "service":
			desc = fmt.Sprintf("Oracle 服务名: %s:%s/%s", host, port, serviceName)
		case "sid":
			desc = fmt.Sprintf("Oracle SID: %s:%s/%s", host, port, serviceName)
		case "tns":
			desc = fmt.Sprintf("Oracle TNS: %s", tnsConnection)
		default:
			return "错误: 不支持的 Oracle 连接类型"
		}
	default:
		return "错误: 不支持的数据库类型"
	}

	return fmt.Sprintf("数据库连接测试成功!\n%s\n用户: %s", desc, username)
}

// connectDatabase 根据类型建立实际的数据库连接 (MySQL / Oracle)
func connectDatabase(dbType, host, port, username, password, connectionType, serviceName, tnsConnection string) (*sql.DB, error) {
	switch strings.ToLower(strings.TrimSpace(dbType)) {
	case "mysql":
		// 对于 MySQL，serviceName 即数据库名
		dbName := strings.TrimSpace(serviceName)
		if dbName == "" {
			return nil, fmt.Errorf("MySQL 需要提供数据库名")
		}

		dsn := fmt.Sprintf("%s:%s@tcp(%s:%s)/%s?charset=utf8mb4&parseTime=True&loc=Local",
			username, password, host, port, dbName)

		db, err := sql.Open("mysql", dsn)
		if err != nil {
			return nil, err
		}
		if err := db.Ping(); err != nil {
			db.Close()
			return nil, err
		}
		return db, nil

	case "oracle":
		var dsn string

		switch connectionType {
		case "service":
			// oracle://user:pass@host:port/serviceName
			dsn = fmt.Sprintf("oracle://%s:%s@%s:%s/%s", username, password, host, port, serviceName)
		case "sid":
			// SID 方式，go-ora 也支持直接使用 /SID
			dsn = fmt.Sprintf("oracle://%s:%s@%s:%s/%s", username, password, host, port, serviceName)
		case "tns":
			// TNS 方式，直接使用 TNS 字符串
			// 注意: 这里假设 tnsConnection 是合法的 TNS 描述
			dsn = fmt.Sprintf("oracle://%s:%s@%s", username, password, tnsConnection)
		default:
			return nil, fmt.Errorf("不支持的 Oracle 连接类型: %s", connectionType)
		}

		db, err := sql.Open("oracle", dsn)
		fmt.Println(dsn)
		if err != nil {
			return nil, err
		}
		if err := db.Ping(); err != nil {
			db.Close()
			return nil, err
		}
		return db, nil
	default:
		return nil, fmt.Errorf("不支持的数据库类型: %s", dbType)
	}
}

func main() {
	// Create an instance of the app structure
	app := NewApp()

	// Create application with options
	err := wails.Run(&options.App{
		Title:  "Excel导入工具",
		Width:  1300,
		Height: 850,
		AssetServer: &assetserver.Options{
			Assets: assets,
		},
		BackgroundColour: &options.RGBA{R: 27, G: 38, B: 54, A: 1},
		OnStartup:        app.startup,
		OnDomReady:       app.domReady,
		OnBeforeClose:    app.beforeClose,
		OnShutdown:       app.shutdown,
		Bind: []interface{}{
			app,
		},
	})

	if err != nil {
		log.Fatal(err)
	}
}
