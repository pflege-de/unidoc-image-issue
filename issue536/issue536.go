package main

import (
	"bytes"
	"encoding/json"
	"errors"
	"fmt"
	"image/png"
	"log"
	"os"
	"strings"

	cfg "github.com/ardanlabs/conf/v3"
	"github.com/boombuler/barcode"
	"github.com/boombuler/barcode/code128"
	"github.com/boombuler/barcode/qr"
	"github.com/unidoc/unioffice/common"
	"github.com/unidoc/unioffice/common/license"
	"github.com/unidoc/unioffice/document"
	"github.com/unidoc/unioffice/document/convert"
	"github.com/unidoc/unioffice/measurement"
	pdflicense "github.com/unidoc/unipdf/v3/common/license"
)

const (
	barcodeWidth      = 3.88
	barcodeHeight     = 0.74
	qrCodeWidthHeight = 1.4
	// {barcode}: 0,74 x 3,88 cm, chose 50x250 to keep dimensions
	barcodeWidthDimension  = 300
	barcodeHeightDimension = 50
	// {qrcode}: 1,4 x 1,4 cm, chose 100x100 to keep dimensions
	qrWidthDimension  = 100
	qrHeightDimension = 100
)

type config struct {
	UniofficeLicenseKey   string `conf:"flag:license,env:LICENSE_KEY"`
	UniofficeCustomerName string `conf:"flag:name,env:CUSTOMER_NAME"`
	UniofficeApiKey       string `conf:"flag:key,env:API_KEY"`
}

func main() {
	var conf config
	txt, err := cfg.Parse("", &conf)
	if err == cfg.ErrHelpWanted {
		fmt.Println(txt)
		os.Exit(0)
	}
	if err != nil {
		fmt.Println(err)
		fmt.Println(txt)
		os.Exit(1)
	}

	switch {
	case conf.UniofficeApiKey != "":
		if err := license.SetMeteredKey(conf.UniofficeApiKey); err != nil {
			fmt.Println(err, "set unioffice api key")
			os.Exit(1)
		}
		if err := pdflicense.SetMeteredKey(conf.UniofficeApiKey); err != nil {
			fmt.Println(err, "set unipdf api key")
			os.Exit(1)
		}
	case conf.UniofficeLicenseKey != "":
		if conf.UniofficeCustomerName == "" {
			fmt.Println("customer name required for license key")
			os.Exit(1)
		}
		if err := license.SetLicenseKey(conf.UniofficeLicenseKey, conf.UniofficeCustomerName); err != nil {
			fmt.Println(err, "set unioffice license key")
			os.Exit(1)
		}
		if err := pdflicense.SetLicenseKey(conf.UniofficeLicenseKey, conf.UniofficeCustomerName); err != nil {
			fmt.Println(err, "set unipdf license key")
			os.Exit(1)
		}
	default:
		fmt.Println("neither api or license key provided")
		os.Exit(1)
	}

	doc, err := document.Open("document.docx")
	if err != nil {
		log.Fatal(err)
	}

	defer doc.Close()

	mappings := make(map[string]string)
	f, err := os.Open("mappings.json")
	if err != nil {
		log.Fatal(err)
	}
	defer f.Close()

	d := json.NewDecoder(f)
	if err := d.Decode(&mappings); err != nil {
		log.Fatal(err)
	}

	log.Printf("%v\n", doc.MergeFields())
	doc.MailMerge(mappings)

	fields := doc.FormFields()

	for _, field := range fields {
		log.Printf("DocField %s[%s]: %v\n", field.Name(), field.Type().String(), field.PossibleValues())
		if field.Type() == document.FormFieldTypeCheckBox {
			// name can be set in word via right click on the checkbox, and setting a value in "bookmark"
			// value is either "true" or "false" for checkboxes
			val, ok := mappings[field.Name()]
			isChecked := ok && strings.ToLower(val) == "true"
			field.SetChecked(isChecked)
		}
	}

	err = fillMappings(doc, mappings)
	if err != nil {
		log.Fatal(err)
	}

	// doc has to be copied so the eventually added images of barcodes are also exported to the PDF
	renewedDoc, err := doc.Copy()
	if err != nil {
		log.Fatal(err)
	}

	temporaryDocxFile, err := os.CreateTemp(".", "*.docx")
	if err != nil {
		log.Fatal(err)
	}
	defer os.Remove(temporaryDocxFile.Name())
	defer temporaryDocxFile.Close()

	err = renewedDoc.SaveToFile(temporaryDocxFile.Name())
	if err != nil {
		log.Fatal(err)
	}
	defer renewedDoc.Close()

	completed, err := document.Open(temporaryDocxFile.Name())
	if err != nil {
		log.Fatal(err)
	}
	defer completed.Close()

	pdfDoc := convert.ConvertToPdf(completed)
	err = pdfDoc.WriteToFile("issue536.pdf")
	if err != nil {
		log.Fatal(err)
	}
}

func fillMappings(doc *document.Document, mappings map[string]string) error {
	doc.StructuredDocumentTags()
	paragraphs := make([]document.Paragraph, len(doc.Paragraphs()))
	for i, p := range doc.Paragraphs() {
		paragraphs[i] = p
	}

	for _, footer := range doc.Footers() {
		for _, paragraph := range footer.Paragraphs() {
			err := replaceBarcodeAndQRCode(paragraph, doc, mappings)
			if err != nil {
				return err
			}
		}
	}

	for _, paragraph := range paragraphs {
		err := replaceBarcodeAndQRCode(paragraph, doc, mappings)
		if err != nil {
			return err
		}
	}

	return nil
}

// replace barcode and qrcode placeholders in a paragraph.
func replaceBarcodeAndQRCode(paragraph document.Paragraph, doc *document.Document, mappings map[string]string) error {
	const openTag = '{'
	const closingTag = '}'

	placeholders := extractPlaceholders(paragraph.Runs(), openTag, closingTag)

	// The first run is replaced with a barcode or qrcode.
	// All other runs associated to the placeholder are deleted.
	for name, runs := range placeholders {
		replaceMe := runs[0]
		deleteMe := runs[1:]

		val, ok := mappings[name]
		if !ok || len(val) < 1 {
			continue
		}

		err := handleRun(replaceMe, name, mappings, doc)
		if err != nil {
			return err
		}
		for _, r := range deleteMe {
			r.Clear()
		}
	}

	return nil
}

// extractTemplates takes a set of runs and extracts the content between openingTag and closingTag.
// The returned map consists of all runs associated with this placeholder.
func extractPlaceholders(runs []document.Run, openingTag, closingTag rune) map[string][]document.Run {
	placeholders := make(map[string][]document.Run)

	var associatedRuns []document.Run
	var constructed string
	var opened bool
	for _, r := range runs {
		for _, c := range r.Text() {
			switch c {
			case openingTag:
				opened = true
				constructed += string(c)
			case closingTag:
				if opened {
					constructed += string(c)
					associatedRuns = append(associatedRuns, r)

					// The placholder name without open & closing tags.
					name := strings.ToLower(constructed[1 : len(constructed)-1])
					placeholders[name] = associatedRuns

					// closing
					opened = false
					associatedRuns = []document.Run{}
					constructed = ""
				}
			default:
				if opened {
					constructed += string(c)
				}
			}
		}

		if opened {
			associatedRuns = append(associatedRuns, r)
		}
	}

	return placeholders
}

func handleRun(r document.Run, key string, mappings map[string]string, doc *document.Document) error {
	// verify if the key is a valid barcode or qrcode placeholder.
	if !(isBarcode(key) || isQRCode(key)) {
		return fmt.Errorf("invalid placeholder detected: [%s]", key)
	}

	replaceValue, ok := mappings[key]
	if !ok {
		return fmt.Errorf("failed to replace [%s], seems like it's missing in the payload. Using key as fallback value", key)
	}

	codeImg, err := insertCode(key, replaceValue)
	if err != nil {
		return err
	}

	err = addImageToDoc(doc, r, codeImg, key)
	if err != nil {
		return err
	}

	return nil
}

func addImageToDoc(doc *document.Document, r document.Run, qrCodeImg barcode.Barcode, key string) error {
	var width, height float64

	if isQRCode(key) {
		width = qrCodeWidthHeight
		height = qrCodeWidthHeight
	} else if isBarcode(key) {
		width = barcodeWidth
		height = barcodeHeight
	} else {
		return errors.New("unsupported code as input")
	}

	buf := new(bytes.Buffer)
	err := png.Encode(buf, qrCodeImg)
	if err != nil {
		return err
	}

	uniImg, err := common.ImageFromBytes(buf.Bytes())
	if err != nil {
		return err
	}

	imgRef, err := doc.AddImage(uniImg)
	if err != nil {
		return err
	}

	err = replaceWithImage(r, imgRef, measurement.Distance(width), measurement.Distance(height))
	if err != nil {
		return err
	}

	return nil
}

func replaceWithImage(r document.Run, imgRef common.ImageRef, width, height measurement.Distance) error {
	r.Clear()
	inlineDrawing, err := r.AddDrawingInline(imgRef)
	if err != nil {
		return err
	}

	inlineDrawing.SetSize(width*measurement.Centimeter, height*measurement.Centimeter)
	return nil
}

func insertCode(key string, replaceValue string) (barcode.Barcode, error) {
	var width, height int
	var code barcode.Barcode
	var err error

	if isQRCode(key) {
		// generated qrcode from value has to be converted to an image to retrieve the bytes
		// bytes are used to create image and the needed image reference by adding it to the document
		code, err = qr.Encode(replaceValue, qr.M, qr.Auto)
		if err != nil {
			return nil, err
		}
		width = qrWidthDimension
		height = qrHeightDimension
	} else if isBarcode(key) {
		code, err = code128.Encode(replaceValue)
		if err != nil {
			return nil, err
		}
		width = barcodeWidthDimension
		height = barcodeHeightDimension
	} else {
		return nil, errors.New("unsupported code as input")
	}

	return barcode.Scale(code, width, height)
}

func trimSpaceAndToLower(str string) string {
	str = strings.TrimSpace(str)
	str = strings.ToLower(str)
	return str
}

// isQRCode returns true in case the string is prefixed by `qrcode`
func isQRCode(str string) bool {
	return strings.HasPrefix(trimSpaceAndToLower(str), "qrcode")
}

// isBarcode returns true in case the string is prefixed by `barcode`
func isBarcode(str string) bool {
	return strings.HasPrefix(trimSpaceAndToLower(str), "barcode")
}
