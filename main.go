package main

import (
	"fmt"
	"os"

	cfg "github.com/ardanlabs/conf/v3"
	"github.com/unidoc/unioffice/common/license"
	"github.com/unidoc/unioffice/document"
	"github.com/unidoc/unioffice/document/convert"
	pdflicense "github.com/unidoc/unipdf/v3/common/license"
)

type config struct {
	UniofficeLicenseKey   string `conf:"flag:license,env:LICENSE_KEY"`
	UniofficeCustomerName string `conf:"flag:name,env:CUSTOMER_NAME"`
	UniofficeApiKey       string `conf:"flag:key,env:API_KEY"`
}

func main() {
	var conf config
	txt, err := cfg.Parse("sample", &conf)
	if err == cfg.ErrHelpWanted {
		fmt.Println(txt)
		os.Exit(0)
	}
	if err != nil {
		fmt.Println(err)
		fmt.Println(txt)
		os.Exit(1)
	}

	if conf.UniofficeApiKey == "" && conf.UniofficeLicenseKey == "" {
	}

	// Register!
	switch {
	case conf.UniofficeApiKey != "":
		if err := license.SetMeteredKey(conf.UniofficeApiKey); err != nil {
			fmt.Println(err, "set unioffice api key")
			os.Exit(1)
		}
		if err := pdflicense.SetMeteredKey(conf.UniofficeApiKey); err != nil {
			fmt.Println(err, "set unioffice api key")
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
			fmt.Println(err, "set unioffice license key")
			os.Exit(1)
		}
	default:
		fmt.Println("neither api or license key provided")
		os.Exit(1)
	}

	fn := "./sample.docx"
	doc, err := document.Open(fn)
	if err != nil {
		fmt.Println("cannot open", fn, "got", err)
		os.Exit(1)
	}
	defer func() {
		if err := doc.Close(); err != nil {
			fmt.Println("cannot close", fn, "got", err)
		}
	}()

	unstreamed, err := doc.Copy()
	if err != nil {
		fmt.Println("cannot created unstreamed copy of", fn, "for persisting images, got", err)
		os.Exit(1)
	}
	defer func() {
		if err := unstreamed.Close(); err != nil {
			fmt.Println("cannot close unstreamed version of", fn, "got", err)
		}
	}()

	pdfDoc := convert.ConvertToPdf(unstreamed)
	on := "./output.pdf"
	if err := pdfDoc.WriteToFile(on); err != nil {
		fmt.Println("cannot store pdf on", on, "got", err)
		os.Exit(1)
	}
}
