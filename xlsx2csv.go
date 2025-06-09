package main

import (
    "encoding/csv"
    "flag"
    "fmt"
    "os"
    "strings"

    "github.com/xuri/excelize/v2"
    "golang.org/x/text/encoding/charmap"
    "golang.org/x/text/encoding/unicode"
    "golang.org/x/text/transform"
)

func getEncoder(encName string) (transform.Transformer, error) {
    switch strings.ToLower(encName) {
    case "utf-8", "utf8":
        return nil, nil // UTF-8 — стандартная кодировка Go
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

func main() {
    xlsxPath := flag.String("xlsx", "", "путь к входному XLSX файлу")
    csvPath := flag.String("csv", "", "путь к выходному CSV файлу")
    encoding := flag.String("enc", "utf-8", "кодировка выходного CSV (utf-8, utf-16le, utf-16be, windows-1251, iso-8859-1)")
    sheetName := flag.String("sheet", "", "название листа для конвертации (по умолчанию первый)")
    delimiter := flag.String("delim", ";", "разделитель полей в CSV")
    flag.Parse()

    if *xlsxPath == "" || *csvPath == "" {
        fmt.Println("Укажите пути к входному XLSX и выходному CSV файлам через -xlsx и -csv")
        return
    }

    f, err := excelize.OpenFile(*xlsxPath)
    if err != nil {
        fmt.Println("Ошибка открытия XLSX:", err)
        return
    }
    defer f.Close()

    var sheet string
    if *sheetName != "" {
        sheet = *sheetName
    } else {
        sheets := f.GetSheetList()
        if len(sheets) == 0 {
            fmt.Println("В файле нет листов")
            return
        }
        sheet = sheets[0]
    }

    rows, err := f.GetRows(sheet)
    if err != nil {
        fmt.Println("Ошибка чтения листа:", err)
        return
    }

    outFile, err := os.Create(*csvPath)
    if err != nil {
        fmt.Println("Ошибка создания CSV:", err)
        return
    }
    defer outFile.Close()

    encoder, err := getEncoder(*encoding)
    if err != nil {
        fmt.Println("Ошибка с кодировкой:", err)
        return
    }

    var writer *csv.Writer
    if encoder != nil {
        // Обернем writer в трансформер для нужной кодировки
        transformWriter := transform.NewWriter(outFile, encoder)
        writer = csv.NewWriter(transformWriter)
    } else {
        writer = csv.NewWriter(outFile)
    }

    // Установка разделителя
    if len(*delimiter) != 1 {
        fmt.Println("Разделитель должен быть один символ")
        return
    }
    writer.Comma = rune((*delimiter)[0])

    for _, row := range rows {
        if err := writer.Write(row); err != nil {
            fmt.Println("Ошибка записи строки CSV:", err)
            return
        }
    }
    writer.Flush()
    if err := writer.Error(); err != nil {
        fmt.Println("Ошибка записи CSV:", err)
        return
    }

    fmt.Println("Конвертация завершена успешно")
}
