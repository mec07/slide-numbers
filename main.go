package main

import (
	"fmt"

	"github.com/unidoc/unioffice/presentation"
)

func main() {
	if err := createExport(); err != nil {
		fmt.Printf("Error generating report: %+v", err)
	}
}

func createExport() error {
	ppt, err := presentation.OpenTemplate("template.pptx")
	if err != nil {
		return err
	}

	for _, s := range ppt.Slides() {
		if err := ppt.RemoveSlide(s); err != nil {
			return err
		}
	}

	layout, err := ppt.GetLayoutByName("1_Table Slide")
	if err != nil {
		return err
	}

	_, err = ppt.AddDefaultSlideWithLayout(layout)
	if err != nil {
		return err
	}

	_, err = ppt.AddDefaultSlideWithLayout(layout)
	if err != nil {
		return err
	}

	if err := ppt.SaveToFile("output.pptx"); err != nil {
		return err
	}
	return nil
}
