package main

import (
	"encoding/csv"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"time"

	"github.com/rs/zerolog"
	"github.com/urfave/cli/v2"
	"github.com/xuri/excelize/v2"
)

var (
	DiffFiles = cli.Command{
		Name:    "process",
		Aliases: []string{"p"},
		Usage:   "Process files in a directory and compare with JSON log",
		Flags: []cli.Flag{
			&cli.StringFlag{
				Name:  "source",
				Usage: "Path to the data JSON file to read",
			},
			&cli.StringFlag{
				Name:  "config",
				Usage: "Path to the config JSON file",
			},
			&cli.StringFlag{
				Name:  "year",
				Usage: "Year for the directory to scan for files",
			},
			&cli.StringFlag{
				Name:  "output",
				Usage: "Path to save the output JSON file",
			},
		},
		Action: func(c *cli.Context) error {
			baseOutputPath := os.Getenv("OUTPUT_DIR")
			offersBaseDir := os.Getenv("OFFERS_BASE_DIR")
			jsonPath := c.String("source")
			configPath := c.String("config")
			yearToSearch := c.String("year")

			outputPath := filepath.Join(baseOutputPath, c.String("output"))
			dirPath := filepath.Clean(offersBaseDir + "\\" + yearToSearch)

			if configPath == "" {
				return cli.Exit("Se requiere el parámetro --config", 1)
			}

			if jsonPath == "" && dirPath == "" {
				return cli.Exit("Se requiere al menos uno de los parámetros --file o --year", 1)
			}

			// Leer y parsear el archivo JSON de configuración
			configData, err := os.ReadFile(configPath)
			if err != nil {
				return cli.Exit(fmt.Sprintf("No se pudo leer el archivo JSON: %v", err), 1)
			}

			var fieldConfigs map[string]FieldConfig
			if err := json.Unmarshal(configData, &fieldConfigs); err != nil {
				return cli.Exit(fmt.Sprintf("No se pudo parsear el archivo JSON: %v", err), 1)
			}

			newFiles, err := compareAndPrompt(jsonPath, dirPath)
			if err != nil {
				return err
			}

			if len(newFiles) > 0 {
				// Process the new files
				var results []Output
				for _, file := range newFiles {
					// Your existing processing function
					output, err := processFile(file, fieldConfigs)
					if err != nil {
						logger.Error().Err(err).Str("file", file).Msg("Problemas al leer el fichero.")
						continue
					}
					results = append(results, output)
				}
				// Save results to JSON
				if err := saveToJSON(results, outputPath); err != nil {
					return err
				}
			}

			return nil
		},
	}
	CountFiles = cli.Command{
		Name:    "count",
		Aliases: []string{"c"},
		Action: func(c *cli.Context) error {
			offersBaseDir := os.Getenv("OFFERS_BASE_DIR")
			dirPath := filepath.Clean(offersBaseDir)
			results, _ := countExcelFiles(dirPath)
			for _, v := range results {
				fmt.Println(v)
			}
			return nil
		},
	}
	CountAndListFiles = cli.Command{
		Name:    "count-and-list",
		Aliases: []string{"cl"},
		Action: func(c *cli.Context) error {
			offersBaseDir := os.Getenv("OFFERS_BASE_DIR")
			dirPath := filepath.Clean(offersBaseDir)
			files, _ := collectExcelFiles(dirPath)
			csvFile, err := os.Create("files.csv")
			if err != nil {
				logger.Fatal().Err(err).Msg("failed creating file")
			}
			defer csvFile.Close()
			csvwriter := csv.NewWriter(csvFile)
			defer csvwriter.Flush()
			var data [][]string
			for _, v := range files {
				row := []string{v}
				data = append(data, row)
			}
			csvwriter.WriteAll(data)
			return nil
		},
	}
	ReadSheetTest = cli.Command{
		Name:    "test",
		Aliases: []string{"t"},
		Flags: []cli.Flag{
			&cli.StringFlag{
				Name:     "file",
				Usage:    "Ruta del archivo Excel",
				Required: true,
			},
		},
		Action: func(c *cli.Context) error {
			fichero, err := excelize.OpenFile(c.String("file"))
			if err != nil {
				logger.Error().Err(err).Msg("error al abrir fichero")
			}
			ficha, err := findSheetByRegex(fichero, "(?i)FICHA")
			if err != nil {
				logger.Error().Err(err).Msg("error al encontrar ficha")
			}

			logger.Info().Msg(ficha)
			return nil
		},
	}
	GetNewFiles = cli.Command{
		Name:    "get",
		Aliases: []string{"g"},
		Action: func(c *cli.Context) error {
			logger.Info().Msg("New function!")
			return nil
		},
	}
	ReadFiles = cli.Command{
		Name:    "read",
		Aliases: []string{"r"},
		Flags: []cli.Flag{
			&cli.StringFlag{
				Name:     "file",
				Usage:    "Ruta del archivo Excel",
				Required: false,
			},
			&cli.StringFlag{
				Name:     "config",
				Usage:    "Ruta del archivo JSON con la configuración de campos",
				Value:    "cell_addressses.json",
				Required: true,
			},
			&cli.StringFlag{
				Name:     "year",
				Usage:    "Año de las ofertas que quieres buscar",
				Required: false,
			},
			&cli.IntFlag{
				Name:     "max",
				Usage:    "Maximo numero de ficheros a analizar",
				Required: false,
			},
			&cli.StringFlag{
				Name:     "output",
				Usage:    "Ruta del archivo de salida JSON",
				Required: false,
			},
			&cli.BoolFlag{
				Name:     "debug",
				Usage:    "Mostrar los contenidos obtenidos de los ficheros",
				Required: false,
			},
		},
		Action: func(c *cli.Context) error {
			baseOutputPath := os.Getenv("OUTPUT_DIR")
			offersBaseDir := os.Getenv("OFFERS_BASE_DIR")

			initTime := time.Now()
			configPath := c.String("config")
			filePath := c.String("file")
			yearToSearch := c.String("year")
			dirPath := filepath.Clean(offersBaseDir + "\\" + yearToSearch)
			maxFiles := c.Int("max")
			outputPath := filepath.Join(baseOutputPath, c.String("output"))
			debugFlag := c.Bool("debug")

			if debugFlag {
				logger = logger.Level(zerolog.DebugLevel)
			}

			if configPath == "" {
				return cli.Exit("Se requiere el parámetro --config", 1)
			}

			if filePath == "" && dirPath == "" {
				return cli.Exit("Se requiere al menos uno de los parámetros --file o --dir", 1)
			}

			// Leer y parsear el archivo JSON de configuración
			configData, err := os.ReadFile(configPath)
			if err != nil {
				return cli.Exit(fmt.Sprintf("No se pudo leer el archivo JSON: %v", err), 1)
			}

			var fieldConfigs map[string]FieldConfig
			if err := json.Unmarshal(configData, &fieldConfigs); err != nil {
				return cli.Exit(fmt.Sprintf("No se pudo parsear el archivo JSON: %v", err), 1)
			}

			if len(c.String("output")) > 0 {
				// crear carpeta si no existe
				if err := ensureDirectory(baseOutputPath); err != nil {
					logger.Fatal().Err(err).Str("baseOutputPath", baseOutputPath).Msg("Problema al crear la carpeta destino del output")
					return err
				}
			}

			var results []Output

			// Procesar un único archivo si se proporciona
			if filePath != "" {
				output, err := processFile(filePath, fieldConfigs)
				if err != nil {
					return cli.Exit(fmt.Sprintf("Error procesando el archivo %s: %v", filePath, err), 1)
				}
				if debugFlag {
					fmt.Println(output)
				}
				results = append(results, output)
			}

			// Recursively process all Excel files in the specified directory
			if yearToSearch != "" {
				excelFiles, err := collectExcelFiles(dirPath)
				if err != nil {
					return cli.Exit(fmt.Sprintf("Failed to read directory: %v", err), 1)
				}

				fmt.Printf("Se encontraron %d archivos Excel en el directorio %s\n", len(excelFiles), dirPath)
				if maxFiles != 0 {
					fmt.Printf("Se analizarán únicamente %d ficheros.", maxFiles)
				}

				for i, path := range excelFiles {
					logger.Info().Int("current_count", i+1).Int("total_count", len(excelFiles)).Str("file_name", filepath.Base(path)).Msg("Procesando archivo...")
					output, err := processFile(path, fieldConfigs)
					if err != nil {
						logger.Error().Str("path", path).Err(err).Msg("Error processing file")
						continue
					}
					if debugFlag {
						fmt.Println(output)
					}
					results = append(results, output)

					if (i + 1) == maxFiles {
						logger.Warn().Int("max_count", maxFiles).Msg("Se ha llegado al numero maximo de ficheros")
						break
					}
				}
			}

			// Escribir resultados acumulados en un archivo JSON
			if len(c.String("output")) > 0 {
				saveToJSON(results, outputPath)
			}
			fmt.Printf("Ejecución completada en %f segundos\n", time.Since(initTime).Seconds())
			return nil
		},
	}
)

func createCLI() *cli.App {
	// retrieve the cli
	app := cli.NewApp()
	app.Name = "Excel Processor"
	app.Usage = "Procesa archivos Excel para extraer datos basados en configuraciones JSON"
	app.Commands = []*cli.Command{
		&CountFiles,
		&CountAndListFiles,
		&ReadFiles,
		&GetNewFiles,
		&DiffFiles,
		&ReadSheetTest,
	}
	return app
}
