package ulla_accueil

import (
	"context"
	"encoding/json"
	"fmt"
	"github.com/GoogleCloudPlatform/functions-framework-go/functions"
	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
	"log"
	"net/http"
	"net/smtp"
	"os"
	"time"
)

func init() {
	functions.HTTP("UllaAccueil", UllaAccueil)
}

func UllaAccueil(w http.ResponseWriter, r *http.Request) {
	link := create()
	w.Header().Set("Location", link)
	w.WriteHeader(http.StatusFound)
}

// Retrieve a token, saves the token, then returns the generated client.
func getClient(config *oauth2.Config) *http.Client {
	tok, _ := tokenFromEnv()
	return config.Client(context.Background(), tok)
}

// Retrieves a token from a local file.
func tokenFromEnv() (*oauth2.Token, error) {
	str := os.Getenv("ACCOUNT_TOKEN")
	tok := &oauth2.Token{}
	err := json.Unmarshal([]byte(str), tok)
	return tok, err
}

func create() string {
	nextMonth := time.Now().AddDate(0, 1, 0)
	month := nextMonth.Month()
	year := nextMonth.Year()

	ctx := context.Background()

	b := os.Getenv("SERVICE_CREDENTIALS")

	config, err := google.ConfigFromJSON([]byte(b), "https://www.googleapis.com/auth/spreadsheets", drive.DriveMetadataScope, drive.DriveFileScope)
	if err != nil {
		log.Fatalf("Unable to parse client secret file to config: %v", err)
	}
	client := getClient(config)

	srv, err := sheets.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		log.Fatalf("Unable to retrieve Sheets client: %v", err)
	}

	drvSrv, err := drive.NewService(ctx, option.WithHTTPClient(client))
	if err != nil {
		log.Fatalf("Unable to retrieve Drive client: %v", err)
	}

	header := fmt.Sprintf("%s %d", getLocalizedMonthName(int(month)), year)
	time1 := "10-12"
	time2 := "14-16"
	time3 := "16-18"

	grey := &sheets.Color{
		Red:   0.8,
		Green: 0.8,
		Blue:  0.8,
	}

	rowdata := []*sheets.RowData{
		{
			Values: []*sheets.CellData{
				{
					UserEnteredValue: &sheets.ExtendedValue{
						StringValue: &header,
					},
					UserEnteredFormat: &sheets.CellFormat{
						TextFormat: &sheets.TextFormat{
							FontSize: 13,
						},
						BackgroundColor: grey,
					},
				},
				{
					UserEnteredValue: &sheets.ExtendedValue{
						StringValue: &time1,
					},
					UserEnteredFormat: &sheets.CellFormat{
						TextFormat: &sheets.TextFormat{
							Bold: true,
						},
						BackgroundColor: grey,
					},
				},
				{
					UserEnteredValue: &sheets.ExtendedValue{
						StringValue: &time2,
					},
					UserEnteredFormat: &sheets.CellFormat{
						TextFormat: &sheets.TextFormat{
							Bold: true,
						},
						BackgroundColor: grey,
					},
				},
				{
					UserEnteredValue: &sheets.ExtendedValue{
						StringValue: &time3,
					},
					UserEnteredFormat: &sheets.CellFormat{
						TextFormat: &sheets.TextFormat{
							Bold: true,
						},
						BackgroundColor: grey,
					},
				},
			},
		},
	}

	for i := 0; i < daysIn(month, year); i++ {
		str := fmt.Sprintf("%d.%d.%d", i+1, month, year)
		rowdata = append(rowdata, &sheets.RowData{
			Values: []*sheets.CellData{
				{
					UserEnteredValue: &sheets.ExtendedValue{
						StringValue: &str,
					},
					UserEnteredFormat: &sheets.CellFormat{
						TextFormat: &sheets.TextFormat{
							Bold:     true,
							FontSize: 10,
						},
						BackgroundColor:     grey,
						HorizontalAlignment: "RIGHT",
					},
				},
				{
					UserEnteredFormat: &sheets.CellFormat{
						Borders: &sheets.Borders{
							Bottom: &sheets.Border{
								Style: "SOLID",
								Width: 1,
							},
							Right: &sheets.Border{
								Style: "SOLID",
								Width: 1,
							},
						},
					},
				},
				{
					UserEnteredFormat: &sheets.CellFormat{
						Borders: &sheets.Borders{
							Bottom: &sheets.Border{
								Style: "SOLID",
								Width: 1,
							},
							Right: &sheets.Border{
								Style: "SOLID",
								Width: 1,
							},
						},
					},
				},
				{
					UserEnteredFormat: &sheets.CellFormat{
						Borders: &sheets.Borders{
							Bottom: &sheets.Border{
								Style: "SOLID",
								Width: 1,
							},
							Right: &sheets.Border{
								Style: "SOLID",
								Width: 1,
							},
						},
					},
				},
			},
		})
	}

	// Create new spreadsheet
	spreadsheet := sheets.Spreadsheet{
		Properties: &sheets.SpreadsheetProperties{
			Title: header,
		},
		Sheets: []*sheets.Sheet{
			{
				Properties: &sheets.SheetProperties{
					Title: "Seite 1",
				},
				Data: []*sheets.GridData{
					{
						RowData: rowdata,
						ColumnMetadata: []*sheets.DimensionProperties{
							{
								PixelSize: 150,
							},
						},
					},
				},
			},
		},
	}

	resp, err := srv.Spreadsheets.Create(&spreadsheet).Do()
	if err != nil {
		log.Fatalf("Unable to create spreadsheet: %v", err)
	}

	fmt.Printf("Created new spreadsheet: %s\n", resp.SpreadsheetId)

	// Create public link
	perm := &drive.Permission{
		Type:               "anyone",
		Role:               "writer",
		AllowFileDiscovery: false,
	}
	_, err = drvSrv.Permissions.Create(resp.SpreadsheetId, perm).Do()
	if err != nil {
		log.Fatalf("Unable to create permission: %v", err)
	}

	publicLink := fmt.Sprintf("https://docs.google.com/spreadsheets/d/%s/edit?usp=sharing", resp.SpreadsheetId)
	fmt.Printf("Created public link: %s", publicLink)

	auth := smtp.PlainAuth("", os.Getenv("EMAIL_SENDER"), os.Getenv("EMAIL_SENDER_PASS"), "smtp.gmail.com")

	// Here we do it all: connect to our server, set up a message and send it

	to := []string{os.Getenv("EMAIL_RECIPIENT")}

	msg := []byte("To: " + os.Getenv("EMAIL_RECIPIENT") + "\r\n" +
		"Subject: Accueil\r\n" +
		"\r\n" +
		"Hier ist der Link für den nächsten Monat\r\n\n" + publicLink)

	err = smtp.SendMail("smtp.gmail.com:587", auth, os.Getenv("EMAIL_SENDER"), to, msg)
	if err != nil {
		log.Fatal(err)
	} else {
		fmt.Println("Email sent!")
	}

	return publicLink
}

func daysIn(m time.Month, year int) int {
	return time.Date(year, m+1, 0, 0, 0, 0, 0, time.UTC).Day()
}

func getLocalizedMonthName(month int) string {
	switch month {
	case 1:
		return "Januar"
	case 2:
		return "Februar"
	case 3:
		return "März"
	case 4:
		return "April"
	case 5:
		return "Mai"
	case 6:
		return "Juni"
	case 7:
		return "Juli"
	case 8:
		return "August"
	case 9:
		return "September"
	case 10:
		return "Oktober"
	case 11:
		return "November"
	case 12:
		return "Dezember"
	default:
		return "Unbekannt"
	}
}
