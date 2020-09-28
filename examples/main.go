package main

import (
	"math/rand"
	"time"

	exceler "github.com/annlumia/go-exceler"
)

func main() {
	r, err := exceler.NewFromFile("./template.xlsx", "Sample")
	if err != nil {
		panic(err)
	}

	data := map[string]interface{}{
		"operator": "Paijo",
		"approver": "Bose Paijo",
		"date":     time.Now().Format("02 January 2006"),
	}

	measurements := make([]interface{}, 24)
	for i := 0; i < len(measurements); i++ {
		measurements[i] = map[string]interface{}{
			"hour":        i,
			"temperature": rand.Float64()*5 + 22,
			"pressure":    rand.Float64()*2 + 3,
			"flowrate":    rand.Float64()*10 + 35,
		}
	}

	data["data"] = measurements

	r.Render(data)
	r.Save("Generated_report.xlsx")
}
