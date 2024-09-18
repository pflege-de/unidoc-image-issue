package main

import (
	"fmt"
	"log"
	"os"

	cfg "github.com/ardanlabs/conf/v3"
	"github.com/unidoc/unioffice/common/license"
	"github.com/unidoc/unioffice/document"
	"github.com/unidoc/unioffice/document/convert"
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

	doc, err := document.Open("crash-test-dummy.docx")
	if err != nil {
		log.Fatal(err)
	}

	defer doc.Close()

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

	convert.RegisterFontsFromDirectory("ttf")

	pdfDoc := convert.ConvertToPdf(completed)
	err = pdfDoc.WriteToFile("crash-test-dummy.pdf")
	if err != nil {
		log.Fatal(err)
	}
}
