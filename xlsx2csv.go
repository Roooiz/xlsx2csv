package main

import (
    "encoding/csv"
    "flag"
    "fmt"
    "io/fs"
    "os"
    "path/filepath"
    "strings"

    "github.com/xuri/excelize/v2"
    "golang.org/x/text/encoding/charmap"
    "golang.org/x/text/encoding/unicode"
    "golang.org/x/text/transform"
)

func getEncoder(encName string) (transform.Transformer, error) {
    switch strings.ToLower(encName) {
    case "utf-8", "utf8":
        return nil, nil
    case "utf-16le":
        return unicode.UTF16(unicode.LittleEndian, unicode.UseBOM).NewEncoder(), nil
    case "utf-16be":
        return unicode.UTF16(unicode.BigEndian, unicode.UseBOM).NewEncoder(), nil
    case "windows-1251":
        return charmap.Windows1251.NewEncoder(), nil
    case "iso-8859-1":
        return charmap.ISO8859_1.NewEncoder(), nil
    default:
        return nil, fmt.Errorf("unsupported encoding: %s", encName)
    }
}

func convertXLSXtoCSV(xlsxPath, csvPath, sheetName, encoding string, delimiter rune) error {
    f, err := excelize.OpenFile(xlsxPath)
    if err != nil {
        return fmt.Errorf("ошибка открытия XLSX %s: %w", xlsxPath, err)
    }
    defer f.Close()

    sheet := sheetName
    if sheet == "" {
        sheets := f.GetSheetList()
        if len(sheets) == 0 {
            return fmt.Errorf("в файле %s нет листов", xlsxPath)
        }
        sheet = sheets[0]
    }

    rows, err := f.GetRows(sheet)
    if err != nil {
        return fmt.Errorf("ошибка чтения листа %s в файле %s: %w", sheet, xlsxPath, err)
    }

    outFile, err := os.Create(csvPath)
    if err != nil {
        return fmt.Errorf("ошибка создания CSV %s: %w", csvPath, err)
    }
    defer outFile.Close()

    encoder, err := getEncoder(encoding)
    if err != nil {
        return fmt.Errorf("ошибка с кодировкой: %w", err)
    }

    var writer *csv.Writer
    if encoder != nil {
        transformWriter := transform.NewWriter(outFile, encoder)
        writer = csv.NewWriter(transformWriter)
    } else {
        writer = csv.NewWriter(outFile)
    }
    writer.Comma = delimiter

    for _, row := range rows {
        if err := writer.Write(row); err != nil {
            return fmt.Errorf("ошибка записи строки CSV: %w", err)
        }
    }
    writer.Flush()
    if err := writer.Error(); err != nil {
        return fmt.Errorf("ошибка записи CSV: %w", err)
    }

    return nil
}

func main() {
    inputPath := flag.String("input", "", "путь к XLSX файлу или папке с XLSX файлами")
    outputDir := flag.String("outdir", "", "путь к папке для сохранения CSV (если вход - папка). Если вход - файл, можно не указывать")
    encoding := flag.String("enc", "utf-8", "кодировка выходного CSV (utf-8, utf-16le, utf-16be, windows-1251, iso-8859-1)")
    sheetName := flag.String("sheet", "", "название листа для конвертации (по умолчанию первый)")
    delimiter := flag.String("delim", ";", "разделитель полей в CSV")
    flag.Parse()

    if *inputPath == "" {
        fmt.Println("Укажите путь к XLSX файлу или папке через -input")
        return
    }
    if len(*delimiter) != 1 {
        fmt.Println("Разделитель должен быть один символ")
        return
    }
    delimRune := rune((*delimiter)[0])

    fi, err := os.Stat(*inputPath)
    if err != nil {
        fmt.Println("Ошибка доступа к пути:", err)
        return
    }

    if fi.IsDir() {
        // Путь к папке — конвертируем все .xlsx файлы
        if *outputDir == "" {
            fmt.Println("Для папки с XLSX укажите папку для CSV через -outdir")
            return
        }
        err := filepath.WalkDir(*inputPath, func(path string, d fs.DirEntry, err error) error {
            if err != nil {
                return err
            }
            if d.IsDir() {
                return nil
            }
            if strings.HasSuffix(strings.ToLower(d.Name()), ".xlsx") {
                csvFileName := strings.TrimSuffix(d.Name(), ".xlsx") + ".csv"
                csvPath := filepath.Join(*outputDir, csvFileName)
                fmt.Printf("Конвертация %s -> %s\n", path, csvPath)
                if err := convertXLSXtoCSV(path, csvPath, *sheetName, *encoding, delimRune); err != nil {
                    fmt.Println("Ошибка:", err)
                }
            }
            return nil
        })
        if err != nil {
            fmt.Println("Ошибка обхода папки:", err)
        }
    } else {
        // Путь к одному файлу
        csvPath := *outputDir
        if csvPath == "" {
            // Если outputDir не указан, создаём CSV рядом с XLSX
            csvPath = strings.TrimSuffix(*inputPath, filepath.Ext(*inputPath)) + ".csv"
        }
        fmt.Printf("Конвертация %s -> %s\n", *inputPath, csvPath)
        if err := convertXLSXtoCSV(*inputPath, csvPath, *sheetName, *encoding, delimRune); err != nil {
            fmt.Println("Ошибка:", err)
        }
    }
}
