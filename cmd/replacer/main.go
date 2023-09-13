package main

import (
	"encoding/json"
	"github.com/harakeishi/gats"
	"github.com/nguyenthenguyen/docx"
	"github.com/santhosh-tekuri/jsonschema/v5"
	"github.com/xuri/excelize/v2"
	"gopkg.in/yaml.v3"
	"log"
	"os"
	"strconv"
	"strings"
)

type Replacer struct {
	data map[string]string
}

func LoadYamlAsJson(filename string) (map[string]interface{}, error) {
	yamlBytes, err := os.ReadFile(filename)
	if err != nil {
		return nil, err
	}

	var data map[string]interface{}
	err = yaml.Unmarshal(yamlBytes, &data)
	if err != nil {
		return nil, err
	}

	return data, nil
}

func LoadYamlAsJsonString(filename string) (string, error) {
	data, err := LoadYamlAsJson(filename)
	if err != nil {
		return "", err
	}

	jsonBytes, err := json.Marshal(data)
	if err != nil {
		return "", err
	}

	return string(jsonBytes), nil
}

func Validate(schemaFilename string, dataFilename string) (*jsonschema.Schema, map[string]interface{}, error) {
	schemaString, err := LoadYamlAsJsonString(schemaFilename)
	data, err := LoadYamlAsJson(dataFilename)

	schema, err := jsonschema.CompileString("schema.json", schemaString)
	if err != nil {
		return nil, nil, err
	}

	return schema, data, schema.Validate(data)
}

func NewReplacer(schemaFilename string, dataFilename string) (*Replacer, error) {
	_, data, err := Validate(schemaFilename, dataFilename)
	if err != nil {
		return nil, err
	}

	d := map[string]string{}
	err = walkData(data, []string{}, func(k, v string) error {
		d["$$"+k+"$$"] = v
		return nil
	})

	return &Replacer{
		data: d,
	}, nil
}

type KeyValueCallback = func(key string, value string) error

func walkData(val interface{}, keys []string, f KeyValueCallback) error {
	switch val := val.(type) {
	case map[string]interface{}:
		for k, v := range val {
			walkData(v, append(keys, k), f)
		}
	case []interface{}:
		for idx, v := range val {
			walkData(v, append(keys, strconv.Itoa(idx)), f)
		}
	default:
		s, err := gats.ToString(val)
		if err != nil {
			return err
		}

		err = f(strings.Join(keys, "."), s)
		if err != nil {
			return err
		}
	}

	return nil
}

func (r *Replacer) replaceDocx(inputFilename string, outputFilename string) error {
	rdoc, err := docx.ReadDocxFile(inputFilename)
	if err != nil {
		return err
	}

	doc := rdoc.Editable()
	if err != nil {
		return err
	}

	for k, v := range r.data {
		doc.Replace(k, v, -1)
	}

	return doc.WriteToFile(outputFilename)
}

func (r *Replacer) replaceXlsx(inputFilename string, outputFilename string) error {
	f, err := excelize.OpenFile(inputFilename)
	if err != nil {
		return err
	}

	defer f.Close()

	sheets := f.GetSheetList()
	for _, sheet := range sheets {
		rows, err := f.GetRows(sheet)
		if err != nil {
			return err
		}
		for ridx, row := range rows {
			for cidx, colCell := range row {
				if val, ok := r.data[colCell]; ok {
					axis, err := excelize.CoordinatesToCellName(cidx+1, ridx+1)
					if err != nil {
						return err
					}
					f.SetCellStr(sheet, axis, val)
				}
			}
		}
	}

	if err := f.SaveAs(outputFilename); err != nil {
		return err
	}

	return nil
}

func main() {
	repl, err := NewReplacer("scheme.yaml", "data.yaml")
	if err != nil {
		log.Fatalln(err)
	}
	err = repl.replaceDocx("test.docx", "out.docx")
	log.Println(err)
	err = repl.replaceXlsx("test.xlsx", "out.xlsx")
	log.Println(err)
}
