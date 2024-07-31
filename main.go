package main

import (
	"bufio"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"time"

	"github.com/djherbis/times"
	"github.com/joho/godotenv"
	"github.com/rs/zerolog"
	"github.com/urfave/cli/v2"
	"github.com/xuri/excelize/v2"
)

// FieldConfig define la estructura para la configuración de los campos
type FieldConfig struct {
	Regex   string `json:"regex"`
	OffsetX int    `json:"offset_x"`
	OffsetY int    `json:"offset_y"`
}

// Output define la estructura del output JSON
type Output struct {
	File         string            `json:"file"`
	BaseFileName string            `json:"base_name"`
	CreatedAt    time.Time         `json:"created_at"`
	ModifiedAt   time.Time         `json:"modified_at"`
	Data         map[string]string `json:"data"`
	UniqueIDs    []string          `json:"unique_ids"`
}

// create a new logger
var logger = zerolog.New(zerolog.ConsoleWriter{Out: os.Stderr, TimeFormat: time.RFC3339}).
	Level(zerolog.InfoLevel).
	With().
	Timestamp().
	Logger()

func main() {
	err := godotenv.Load()
	if err != nil {
		logger.Fatal().Msg("Error loading .env file")
	}

	app := createCLI()

	err = app.Run(os.Args)
	if err != nil {
		logger.Fatal().Err(err).Msg("Error al ejecutar el inicio del programa")
	}
}

func processFile(filePath string, fieldConfigs map[string]FieldConfig) (Output, error) {
	var output Output
	// Abrir el archivo Excel
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return output, fmt.Errorf("no se pudo abrir el archivo Excel: %v", err)
	}
	defer f.Close()

	results := make(map[string]string)
	var uniqueIDs []string

	// Buscar hojas que coincidan con las expresiones regulares
	fichaSheet, err := findSheetByRegex(f, "(?i)^FICHA$")
	if err != nil {
		return output, fmt.Errorf("no se pudo encontrar la hoja de FICHA: %v", err)
	}

	sapSheet, err := findSheetByRegex(f, "(?i)^SAP$")
	if err != nil {
		sapSheet, err = findSheetByRegex(f, "(?i)^OFERTA$")
		if err != nil {
			return output, fmt.Errorf("no se pudo encontrar la hoja de SAP u OFERTA: %v", err)
			// logger.Error().Err(fmt.Errorf("no se pudo encontrar la hoja de SAP: %v", err))
		}
	}

	// Procesar datos de la hoja FICHA
	for field, config := range fieldConfigs {
		// Buscar la celda que coincide con la expresión regular (case insensitive)
		cell, err := findCellByRegex(f, fichaSheet, "(?i)"+config.Regex)
		if err != nil {
			logger.Debug().Msg(fmt.Sprintf("Advertencia: %v. Ignorando el campo %s", err, field))
			continue
		}

		if cell != "" {
			// Obtener las coordenadas de la celda encontrada
			x, y, err := excelize.CellNameToCoordinates(cell)
			if err != nil {
				return output, fmt.Errorf("no se pudo convertir el nombre de la celda a coordenadas: %v", err)
			}

			// Aplicar los offsets para obtener la celda deseada
			targetCell, err := excelize.CoordinatesToCellName(x+config.OffsetX, y-config.OffsetY)
			if err != nil {
				return output, fmt.Errorf("no se pudo convertir las coordenadas a nombre de celda: %v", err)
			}

			// Obtener el valor de la celda objetivo
			value, err := f.GetCellValue(fichaSheet, targetCell)
			if err != nil {
				return output, fmt.Errorf("no se pudo obtener el valor de la celda %s: %v", targetCell, err)
			}

			results[field] = value
		}
	}

	// Obtener datos de ultima modificacion y creacion
	finfo, err := os.Stat(filePath)
	if err != nil {
		logger.Error().Err(err).Msg("No se pudo obtener info del archivo Excel")
	}

	// Crear la estructura de output
	output = Output{
		File:         filePath,
		BaseFileName: finfo.Name(),
		CreatedAt:    times.Get(finfo).BirthTime(),
		ModifiedAt:   finfo.ModTime(),
		Data:         results,
	}

	// Procesar datos de la hoja SAP
	uniqueIDs, err = findUniqueValues(f, sapSheet, "(?i)Registral")
	if err != nil {
		return output, fmt.Errorf("no se pudo obtener los valores únicos de la columna Registral: %v", err)
	}

	output.UniqueIDs = uniqueIDs

	return output, nil
}

// findSheetByRegex busca una hoja en el archivo que coincida con la expresión regular
func findSheetByRegex(f *excelize.File, regex string) (string, error) {
	re, err := regexp.Compile(regex)
	if err != nil {
		return "", err
	}

	for _, sheet := range f.GetSheetList() {
		if re.MatchString(sheet) {
			return sheet, nil
		}
	}

	return "", fmt.Errorf("no se encontró ninguna hoja que coincida con la expresión regular: %s", regex)
}

// findCellByRegex busca una celda en la hoja que coincida con la expresión regular
func findCellByRegex(f *excelize.File, sheetName, regex string) (string, error) {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return "", err
	}

	re, err := regexp.Compile(regex)
	if err != nil {
		return "", err
	}

	for rowIndex, row := range rows {
		for colIndex, cell := range row {
			if re.MatchString(cell) {
				return excelize.CoordinatesToCellName(colIndex+1, rowIndex+1)
			}
		}
	}

	return "", fmt.Errorf("no se encontró ninguna celda que coincida con la expresión regular: %s", regex)
}

// findUniqueValues busca valores únicos en una columna que coincide con la expresión regular en la hoja
func findUniqueValues(f *excelize.File, sheetName, regex string) ([]string, error) {
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, err
	}

	re, err := regexp.Compile(regex)
	if err != nil {
		return nil, err
	}

	var columnIndex int = -1
	var uniqueValues []string
	valueSet := make(map[string]struct{})

	// Encontrar el índice de la columna que coincide con la expresión regular
	for colIndex, cell := range rows[0] {
		if re.MatchString(cell) {
			columnIndex = colIndex
			break
		}
	}

	if columnIndex == -1 {
		return nil, fmt.Errorf("no se encontró ninguna columna que coincida con la expresión regular: %s", regex)
	}

	// Recoger valores únicos de la columna
	for _, row := range rows[1:] {
		if len(row) > columnIndex {
			value := strings.TrimSpace(row[columnIndex])
			if value != "" {
				if _, found := valueSet[value]; !found {
					valueSet[value] = struct{}{}
					uniqueValues = append(uniqueValues, value)
				}
			}
		}
	}

	return uniqueValues, nil
}

// collectExcelFiles recursively finds all .xlsx files in the provided directory.
func countExcelFiles(root string) ([]string, error) {
	var results []string
	var x_files int
	var m_files int
	var s_files int
	var d_files int
	logger.Info().Str("root folder", root).Msg("Analizando carpeta raíz en busca de ficheros Excel...")
	err := filepath.Walk(root, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if !info.IsDir() && !strings.HasPrefix(info.Name(), "~") {
			switch filepath.Ext(info.Name()) {
			case ".xlsx":
				x_files += 1
			case ".xlsm":
				m_files += 1
			case ".xls":
				s_files += 1
			default:
				d_files += 1
			}
		}
		return nil
	})
	results = append(results, fmt.Sprintf("xlsx files total count: %d", x_files))
	results = append(results, fmt.Sprintf("xlsm files total count: %d", m_files))
	results = append(results, fmt.Sprintf("xls files total count: %d", s_files))
	results = append(results, fmt.Sprintf("rest of the files total count: %d", d_files))
	return results, err
}

// collectExcelFiles recursively finds all .xlsx files in the provided directory.
func collectExcelFiles(root string) ([]string, error) {
	var files []string
	logger.Info().Str("root folder", root).Msg("Analizando carpeta raíz en busca de ficheros Excel...")
	err := filepath.Walk(root, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}
		if !info.IsDir() && strings.HasSuffix(info.Name(), ".xlsx") && !strings.HasPrefix(info.Name(), "~") {
			files = append(files, path)
		}
		return nil
	})
	return files, err
}

// Function to ensure directory exists, and create if not
func ensureDirectory(dir string) error {
	if _, err := os.Stat(dir); os.IsNotExist(err) {
		logger.Warn().Str("outputpath", dir).Msg("Carpeta no encontrada. Creando...")
		err = os.MkdirAll(dir, 0755)
		if err != nil {
			return err
		}
		logger.Info().Str("outputpath", dir).Msg("Carpeta creada con éxito.")
	}
	return nil
}

func saveToJSON(data []Output, path string) error {
	jsonOutput, err := json.MarshalIndent(data, "", "  ")
	if err != nil {
		return cli.Exit(fmt.Sprintf("Error al generar JSON: %v", err), 1)
	}
	logger.Info().Msg("Creando fichero JSON...")
	if err := os.WriteFile(path, jsonOutput, 0644); err != nil {
		return cli.Exit(fmt.Sprintf("Error al escribir el archivo JSON: %v", err), 1)
	}
	fmt.Printf("Resultados guardados en el archivo: %s\n", path)
	return nil
}

// Function to read JSON file and return the list of processed files
func readProcessedFiles(jsonPath string) (map[string]Output, error) {
	var outputs []Output
	fileData, err := os.ReadFile(jsonPath)
	if err != nil {
		return nil, err
	}
	if err := json.Unmarshal(fileData, &outputs); err != nil {
		return nil, err
	}

	processedFiles := make(map[string]Output)
	for _, output := range outputs {
		processedFiles[output.File] = output
	}
	return processedFiles, nil
}

// Function to compare files and prompt user for permission
func compareAndPrompt(jsonPath, dirPath string) ([]string, error) {
	// Read processed files from JSON
	processedFiles, err := readProcessedFiles(jsonPath)
	if err != nil {
		return nil, err
	}

	// Collect all files from directory
	allFiles, err := collectExcelFiles(dirPath)
	if err != nil {
		return nil, err
	}

	// Find files not in processedFiles
	var newFiles []string
	for _, file := range allFiles {
		if _, found := processedFiles[file]; !found {
			newFiles = append(newFiles, file)
		}
	}

	// Prompt user for permission to process new files
	if len(newFiles) > 0 {
		fmt.Println("Nuevos archivos encontrados:")
		for _, file := range newFiles {
			fmt.Println(file)
		}

		fmt.Print("¿Desea procesar estos archivos? (s/n): ")
		reader := bufio.NewReader(os.Stdin)
		response, _ := reader.ReadString('\n')
		response = strings.ToLower(strings.Replace(response, "\r\n", "", -1))
		fmt.Println(response)
		if response != "s" && response != "y" {
			return nil, fmt.Errorf("el usuario decidió no procesar los nuevos archivos")
		}
	} else {
		logger.Info().Msg("No hay nuevos ficheros!")
	}

	return newFiles, nil
}
