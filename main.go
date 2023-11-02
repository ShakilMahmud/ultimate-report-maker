package main

import (
	"database/sql"
	"fmt"
	"net/http"
	"os"
	"path/filepath"

	"github.com/gin-contrib/cors"
	"github.com/gin-gonic/gin"
	_ "github.com/go-sql-driver/mysql"
	"github.com/tealeg/xlsx"
)

func main() {
	r := gin.Default()

	// Configure CORS middleware
	config := cors.DefaultConfig()
	config.AllowOrigins = []string{"*"}
	config.AllowMethods = []string{"GET", "POST", "OPTIONS"}
	r.Use(cors.New(config))

	// Test
	r.GET("/", HelloFromApp)

	// Define an API route to handle the database query and Excel file generation.
	r.POST("/generate-excel", GenerateExcel)

	// Run the server on port 8080.
	port := os.Getenv("PORT") // Get the port from the environment variable
	if port == "" {
		port = "8080" // Default to 8080 if PORT environment variable is not set
	}
	r.Run(":" + port)
}
func HelloFromApp(c *gin.Context) {
	c.JSON(http.StatusAccepted, gin.H{"Test": "Hello From App"})
}

func GenerateExcel(c *gin.Context) {
	// Parse JSON request body to get database credentials and query.
	var requestBody struct {
		DBUser     string `json:"db_user"`
		DBPassword string `json:"db_password"`
		DBHost     string `json:"db_host"`
		DBName     string `json:"db_name"`
		Query      string `json:"query"`
	}
	if err := c.BindJSON(&requestBody); err != nil {
		c.JSON(http.StatusBadRequest, gin.H{"error": "Invalid JSON data"})
		return
	}

	// Create a MySQL DSN (Data Source Name).
	dsn := fmt.Sprintf("%s:%s@tcp(%s)/%s", requestBody.DBUser, requestBody.DBPassword, requestBody.DBHost, requestBody.DBName)

	// Open a database connection.
	db, err := sql.Open("mysql", dsn)
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to connect to the database"})
		return
	}
	defer db.Close()

	// Execute the MySQL query.
	rows, err := db.Query(requestBody.Query)
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to execute the query"})
		return
	}
	defer rows.Close()

	// Create an Excel file and a worksheet.
	file := xlsx.NewFile()
	sheet, err := file.AddSheet("Sheet1")
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to create Excel sheet"})
		return
	}

	// Fetch column names.
	columns, err := rows.Columns()
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to fetch column names"})
		return
	}

	// Add column names to the first row.
	headerRow := sheet.AddRow()
	for _, col := range columns {
		cell := headerRow.AddCell()
		cell.SetString(col)
	}

	// Fetch and add rows to the worksheet.
	for rows.Next() {
		// Create new slices for each row iteration.
		values := make([]interface{}, len(columns))
		valuePtrs := make([]interface{}, len(columns))
		for i := range columns {
			valuePtrs[i] = &values[i]
		}

		err := rows.Scan(valuePtrs...)
		if err != nil {
			c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to scan row values"})
			return
		}

		row := sheet.AddRow()
		for _, val := range values {
			cell := row.AddCell()
			// Handle type conversion based on the actual data types in your database.
			switch v := val.(type) {
			case int64:
				cell.SetInt64(v)
			case float64:
				cell.SetFloatWithFormat(v, "0.00") // Specify the desired format.
			case []byte:
				// Convert []byte to string assuming it's UTF-8 encoded data.
				cell.SetString(string(v))
			default:
				cell.SetString(fmt.Sprintf("%v", v))
			}
		}
	}

	// Generate a unique filename for the Excel file.
	excelFileName := fmt.Sprintf("result_%s.xlsx", requestBody.DBName)

	// Save the Excel file to a temporary location.
	tempDir := os.TempDir()
	tempFilePath := filepath.Join(tempDir, excelFileName)
	err = file.Save(tempFilePath)
	if err != nil {
		c.JSON(http.StatusInternalServerError, gin.H{"error": "Failed to save Excel file"})
		return
	}

	// Provide the Excel file for download.
	c.Header("Content-Description", "File Transfer")
	c.Header("Content-Disposition", fmt.Sprintf("attachment; filename=%s", excelFileName))
	c.Header("Content-Type", "application/octet-stream")

	// Serve the file for download.
	c.File(tempFilePath)

	// Clean up: Delete the temporary Excel file after it's been served.
	defer os.Remove(tempFilePath)

}
