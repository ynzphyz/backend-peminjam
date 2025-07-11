package main

import (
	"bytes"
	"context"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/rs/cors"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/docs/v1"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

	type FormData struct {
		Nama                string
		Kelas               string
		NIS                 string
		NoWA                string
		NamaAlat            string
		JumlahAlat          int
		TanggalPinjam       string
		TanggalKembali      string
		Keterangan          string
		KeteranganPinjam    string // New field for original loan description
		FotoPath            string
		PeminjamanFotoPath  string // New field for peminjaman photo URL
		ApprovalStatus      string // New field for approval status: Pending, Approved, Rejected
		KondisiAlat         string // Added for pengembalian kondisi alat
		KeteranganPengembalian string // Added for pengembalian keterangan

		ApprovalDate        string // New field for approval date
		ApproverName        string // New field for approver name
	}

func getServices() (*sheets.Service, *drive.Service, *docs.Service, error) {
	b, err := os.ReadFile("credentials.json")
	if err != nil {
		return nil, nil, nil, fmt.Errorf("unable to read credentials: %v", err)
	}
	config, err := google.ConfigFromJSON(b, sheets.SpreadsheetsScope, drive.DriveFileScope, docs.DocumentsScope)
	if err != nil {
		return nil, nil, nil, fmt.Errorf("unable to parse credentials: %v", err)
	}
	client := getClient(config)

	sheetsService, _ := sheets.NewService(context.Background(), option.WithHTTPClient(client))
	driveService, _ := drive.NewService(context.Background(), option.WithHTTPClient(client))
	docsService, _ := docs.NewService(context.Background(), option.WithHTTPClient(client))

	return sheetsService, driveService, docsService, nil
}

func parseTanggal(t string) time.Time {
	d, _ := time.Parse("2006-01-02", t)
	return d
}

func saveFileLocally(file io.Reader, filename string) (string, error) {
	os.MkdirAll("uploads", os.ModePerm)
	path := filepath.Join("uploads", fmt.Sprintf("%d_%s", time.Now().UnixNano(), filename))
	f, err := os.Create(path)
	if err != nil {
		return "", err
	}
	defer f.Close()
	io.Copy(f, file)
	return path, nil
}

func uploadToDrive(localPath, filename string, driveService *drive.Service) (string, error) {
	f, _ := os.Open(localPath)
	defer f.Close()

	meta := &drive.File{
		Name:     filename,
		Parents:  []string{"19iloK_NHLVzAhy_I_dt6RH6aNRaTQkAV"},
		MimeType: "image/jpeg",
	}
	file, err := driveService.Files.Create(meta).Media(f).Do()
	if err != nil {
		return "", err
	}

	driveService.Permissions.Create(file.Id, &drive.Permission{Role: "reader", Type: "anyone"}).Do()
	return fmt.Sprintf("https://drive.google.com/uc?id=%s", file.Id), nil
}

func generateSurat(form FormData, nomorUrut int, driveService *drive.Service, docsService *docs.Service) (pdfURL, docURL string, err error) {
	templateID := "1RK2I4oAUvPFTlv98Hp5bDlassulBFvrASuhs5-riVUM"
	pdfFolder := "1HhZncgqeqEzgTkMQZOBC9HAsPTIB0zTv"
	title := fmt.Sprintf("Formulir Peminjaman %04d - %s", nomorUrut, form.Nama)

	copy, err := driveService.Files.Copy(templateID, &drive.File{Name: title}).Do()
	if err != nil {
		return "", "", fmt.Errorf("failed to copy template: %v", err)
	}
	docID := copy.Id
	docURL = fmt.Sprintf("https://docs.google.com/document/d/%s/edit", docID)

	docFolder := "1Y3cvxCOy4M0GtRPe7A1DrAg1iji5O0lQ"
	
	_, err = driveService.Files.Update(docID, nil).
		AddParents(docFolder).
		RemoveParents("root").
		Do()
	if err != nil {
		log.Println("⚠️ Gagal memindahkan file ke folder Dokumen:", err)
	}

	replacements := map[string]string{
		"<<NMR>>":    fmt.Sprintf("%04d", nomorUrut),
		"<<TGL>>":    time.Now().Format("02 January 2006"),
		"<<NAMA>>":   form.Nama,
		"<<KLS>>":    form.Kelas,
		"<<NIS>>":    form.NIS,
		"<<NO>>":     form.NoWA,
		"<<NMALT>>":  form.NamaAlat,
		"<<JML>>":    fmt.Sprintf("%d", form.JumlahAlat),
		"<<TGLPMJ>>": form.TanggalPinjam,
		"<<TGLPGN>>": form.TanggalKembali,
		"<<LMPJM>>":  fmt.Sprintf("%d hari", int(parseTanggal(form.TanggalKembali).Sub(parseTanggal(form.TanggalPinjam)).Hours()/24)),
		"<<KET>>":    form.Keterangan,
	}

	var reqs []*docs.Request
	for key, val := range replacements {
		log.Printf("DEBUG: Replacement key: '%s', value: '%s'\n", key, val)
		reqs = append(reqs, &docs.Request{
			ReplaceAllText: &docs.ReplaceAllTextRequest{
				ContainsText: &docs.SubstringMatchCriteria{Text: key, MatchCase: true},
				ReplaceText:  val,
			},
		})
	}
	respBatchUpdate, err := docsService.Documents.BatchUpdate(docID, &docs.BatchUpdateDocumentRequest{Requests: reqs}).Do()
	if err != nil {
		log.Printf("ERROR: BatchUpdate failed: %v\n", err)
		// Cannot use http.Error here because w is not in scope in this function
		// Just log the error and return
		return
	} else {
		log.Printf("INFO: BatchUpdate response: %+v\n", respBatchUpdate)
	}

	if form.FotoPath != "" {
		doc, err := docsService.Documents.Get(docID).Do()
		if err == nil {
			var index int64
			for _, c := range doc.Body.Content {
				if c.Paragraph != nil {
					for _, e := range c.Paragraph.Elements {
						if e.TextRun != nil && strings.Contains(e.TextRun.Content, "<<FOTO>>") {
							index = e.StartIndex
							break
						}
					}
				}
			}
			end := index + int64(len("<<FOTO>>"))
			imgReq := []*docs.Request{
				{DeleteContentRange: &docs.DeleteContentRangeRequest{
					Range: &docs.Range{StartIndex: index, EndIndex: end},
				}},
				{InsertInlineImage: &docs.InsertInlineImageRequest{
					Location: &docs.Location{Index: index},
					Uri:      form.FotoPath,
					ObjectSize: &docs.Size{
						Width:  &docs.Dimension{Magnitude: 400, Unit: "PT"},
						Height: &docs.Dimension{Magnitude: 225, Unit: "PT"},
					},
				}},
			}
			docsService.Documents.BatchUpdate(docID, &docs.BatchUpdateDocumentRequest{Requests: imgReq}).Do()
		}
	}

	export, err := driveService.Files.Export(docID, "application/pdf").Download()
	if err != nil {
		return "", "", fmt.Errorf("failed to export PDF: %v", err)
	}
	tmp := filepath.Join("uploads", fmt.Sprintf("%s.pdf", title))
	out, _ := os.Create(tmp)
	io.Copy(out, export.Body)
	out.Close()

	file, _ := os.Open(tmp)
	pdf, _ := driveService.Files.Create(&drive.File{
		Name:     filepath.Base(tmp),
		Parents:  []string{pdfFolder},
		MimeType: "application/pdf",
	}).Media(file).Do()
	file.Close()
	os.Remove(tmp)

	driveService.Permissions.Create(pdf.Id, &drive.Permission{Role: "reader", Type: "anyone"}).Do()
	driveService.Permissions.Create(docID, &drive.Permission{Role: "reader", Type: "anyone"}).Do()

	pdfURL = fmt.Sprintf("https://drive.google.com/uc?id=%s", pdf.Id)
	return pdfURL, docURL, nil
}

func normalizePhoneNumber(no string) string {
	log.Printf("DEBUG: normalizePhoneNumber input: '%s'", no)
	no = strings.TrimSpace(no)
	log.Printf("DEBUG: after TrimSpace: '%s'", no)
	no = strings.ReplaceAll(no, " ", "")
	no = strings.ReplaceAll(no, "-", "")
	no = strings.ReplaceAll(no, "(", "")
	no = strings.ReplaceAll(no, ")", "")
	log.Printf("DEBUG: after removing spaces and symbols: '%s'", no)
	if strings.HasPrefix(no, "+") {
		no = "62" + no[1:]
		log.Printf("DEBUG: after handling '+': '%s'", no)
	} else if strings.HasPrefix(no, "0") {
		no = "62" + no[1:]
		log.Printf("DEBUG: after handling '0': '%s'", no)
	} else if strings.HasPrefix(no, "62") {
		// number already in correct format, do nothing
		log.Printf("DEBUG: number starts with '62', no change")
	} else {
		// invalid format, clear the number
		log.Printf("DEBUG: number invalid format, clearing")
		no = ""
	}
	log.Printf("DEBUG: normalizePhoneNumber output: '%s'", no)
	return no
}

func kirimPesanWaBangkit(no string, pesan string) error {
	no = normalizePhoneNumber(no)
	log.Printf("DEBUG: Nomor WA setelah normalisasi: '%s'\n", no)
	if !strings.HasPrefix(no, "62") {
		return fmt.Errorf("❌ Format nomor WA tidak valid (harus 62...), silakan isi ulang")
	}

	payload := map[string]string{
		"api_key": "tW3CWRv5NyTGKuhsrmcRqoKYEnCMVQ",
		"sender":  "6287760573989",
		"number":  no,
		"message": pesan,
	}
	body, _ := json.Marshal(payload)

	client := &http.Client{Timeout: 10 * time.Second}
	resp, err := client.Post("https://wa.bangkitsolusibangsa.id/send-message", "application/json", bytes.NewBuffer(body))
	if err != nil {
		return err
	}
	defer resp.Body.Close()

	if resp.StatusCode >= 300 {
		return fmt.Errorf("WA API error: %s", resp.Status)
	}
	return nil
}

func getSalam() string {
	hour := time.Now().Hour()
	switch {
	case hour < 11:
		return "Selamat pagi"
	case hour < 15:
		return "Selamat siang"
	case hour < 18:
		return "Selamat sore"
	default:
		return "Selamat malam"
	}
}

func getPeminjamDetailsByID(sheetsService *sheets.Service, sheetId string, nis string) (name string, noWA string, err error) {
	resp, err := sheetsService.Spreadsheets.Values.Get(sheetId, "Form Peminjam!A5:Z").Do()
	if err != nil {
		return "", "", err
	}
	if resp == nil || resp.Values == nil {
		return "", "", fmt.Errorf("data peminjam kosong")
	}
	for _, row := range resp.Values {
		if len(row) > 4 {
			sheetNIS := strings.TrimSpace(fmt.Sprintf("%v", row[4]))
			log.Printf("DEBUG: Checking row NIS: '%s' against form NIS: '%s'\n", sheetNIS, strings.TrimSpace(nis))
			if sheetNIS == strings.TrimSpace(nis) {
				log.Printf("DEBUG: Matching row data: %+v\n", row)
				if len(row) > 2 {
					name = strings.TrimSpace(fmt.Sprintf("%v", row[2]))
				}
				if len(row) > 5 {
					noWA = strings.TrimSpace(fmt.Sprintf("%v", row[5]))
				}
				log.Printf("DEBUG: Found matching NIS. Name: '%s', NoWA: '%s'\n", name, noWA)
				return name, noWA, nil
			}
		}
	}
	log.Printf("DEBUG: NIS '%s' not found in sheet\n", nis)
	return "", "", fmt.Errorf("nis peminjam tidak ditemukan")
}

func handlePinjam(w http.ResponseWriter, r *http.Request) {
	r.ParseMultipartForm(10 << 20)
	jumlah, _ := strconv.Atoi(r.FormValue("jumlahAlat"))
	form := FormData{
		Nama:           r.FormValue("nama"),
		Kelas:          r.FormValue("kelas"),
		NIS:            r.FormValue("nis"),
		NoWA:           r.FormValue("noWa"),
		NamaAlat:       r.FormValue("namaAlat"),
		JumlahAlat:     jumlah,
		TanggalPinjam:  r.FormValue("tanggalPinjam"),
		TanggalKembali: r.FormValue("tanggalKembali"),
		Keterangan:     r.FormValue("keterangan"),
	}

	// Save the uploaded file locally first
	var localPath string
	file, handler, err := r.FormFile("foto")
	if err == nil {
		defer file.Close()
		localPath, _ = saveFileLocally(file, handler.Filename)
	}

	// Respond immediately to the client
	w.Write([]byte("✅ Data berhasil diterima dan sedang diproses"))

	// Fetch sheet data before starting goroutine
	sheetId := "1uULs6gLCAeLVeOI-qjdIcb4pRod-mC6g4Cu9TvtIVak"
	sheetData := func() *sheets.ValueRange {
		sheetsService, _, _, err := getServices()
		if err != nil {
			log.Println("Service error:", err)
			return nil
		}
		resp, err := sheetsService.Spreadsheets.Values.Get(sheetId, "Form Peminjam!B5:B").Do()
		if err != nil {
			log.Println("❌ Gagal mengambil data dari Sheets:", err)
			return nil
		}
		return resp
	}()
	if sheetData == nil {
		log.Println("❌ Tidak dapat mengambil data sheet, melanjutkan tanpa update nama")
	}

	// Process the heavy work asynchronously
	go func(form FormData, localPath string, sheetData *sheets.ValueRange) {
		sheetsService, driveService, docsService, err := getServices()
		if err != nil {
			log.Println("Service error:", err)
			return
		}

		// Upload file to Drive if available
		if localPath != "" {
			url, err := uploadToDrive(localPath, filepath.Base(localPath), driveService)
			if err == nil {
				form.FotoPath = url
				log.Println("✅ Link foto pengembalian:", form.FotoPath)
			} else {
				log.Println("❌ Gagal upload foto pengembalian ke Drive:", err)
			}
			os.Remove(localPath)
		}

			// Fetch peminjam details by ID (assuming form.NIS is the peminjam ID)
		// Gunakan langsung nama dan WA dari form yang baru saja dikirim
		name := form.Nama
		noWA := form.NoWA

		// Fallback jika kosong
		if noWA == "" {
			if sheetData != nil {
				rowToUpdate := len(sheetData.Values) + 4
				if rowToUpdate > 5 {
					rangeGet := fmt.Sprintf("Form Peminjam!F%d", rowToUpdate)
					respNoWA, err := sheetsService.Spreadsheets.Values.Get(sheetId, rangeGet).Do()
					if err == nil && len(respNoWA.Values) > 0 && len(respNoWA.Values[0]) > 0 {
						noWA = strings.TrimSpace(fmt.Sprintf("%v", respNoWA.Values[0][0]))
						log.Println("✅ Fallback: NoWA diambil dari baris terakhir:", noWA)
					}
				}
			}
		}
		form.NoWA = noWA

			if err != nil {
				log.Println("⚠️ Gagal mengambil data peminjam:", err)
				noWA = form.NoWA // fallback to form NoWA if error
			} else {
				log.Printf("DEBUG: Raw NoWA fetched from sheet: '%s'\n", noWA)
				noWA = strings.TrimSpace(noWA)
				if noWA == "" {
					log.Println("⚠️ NoWA dari sheet kosong, menggunakan form.NoWA sebagai fallback")
					noWA = form.NoWA
				}
				if noWA == "" {
					log.Println("⚠️ Nomor WA peminjam kosong, tidak dapat mengirim pesan WA")
				}
				form.NoWA = noWA
				// Use form.Nama if provided, else use sheet name
				if form.Nama == "" {
					form.Nama = name
				} else if form.Nama != name {
					// Update sheet with new name from form
					// Update only the last row (newly added row) to avoid overwriting older rows with same NIS
					if sheetData != nil {
						rowToUpdate := len(sheetData.Values) + 4
						// Prevent updating row 5 (original data)
						if rowToUpdate > 5 {
							writeRange := fmt.Sprintf("Form Peminjam!C%d", rowToUpdate)
							values := [][]interface{}{{form.Nama}}
							vr := &sheets.ValueRange{Values: values}
							_, err := sheetsService.Spreadsheets.Values.Update(sheetId, writeRange, vr).ValueInputOption("USER_ENTERED").Do()
							if err != nil {
								log.Println("⚠️ Gagal update nama di sheet:", err)
							} else {
								log.Println("INFO: Nama di sheet berhasil diperbarui menjadi:", form.Nama)
							}
						} else {
							log.Println("INFO: Tidak memperbarui nama di baris 5 atau sebelumnya")
						}
					} else {
						log.Println("⚠️ Data sheet tidak tersedia, tidak dapat update nama")
					}
				}
			}

			resp, err := sheetsService.Spreadsheets.Values.Get(sheetId, "Form Peminjam!B5:B").Do()
			if err != nil {
				log.Println("❌ Gagal mengambil data dari Sheets:", err)
				return
			}
			log.Printf("DEBUG: Sheets API response: %+v\n", resp)

			var row int
			var pdf, doc string

			if resp == nil || resp.Values == nil || len(resp.Values) == 0 {
				log.Println("❌ Response dari Sheets kosong, memulai dari baris 1")
				row = 1
			} else {
				row = len(resp.Values) + 1
			}

			writeRange := fmt.Sprintf("Form Peminjam!A%d", row+4)

			// Continue processing with row
			pdf, doc, err = generateSurat(form, row, driveService, docsService)
			if err != nil {
				log.Println("❌ Gagal generate surat:", err)
				return
			}

			values := []interface{}{
				fmt.Sprintf("%04d", row), time.Now().Format("2006-01-02"), form.Nama, form.Kelas, form.NIS,
				form.NoWA, form.NamaAlat, form.JumlahAlat, form.TanggalPinjam, form.TanggalKembali,
				form.Keterangan, fmt.Sprintf("%d hari", int(parseTanggal(form.TanggalKembali).Sub(parseTanggal(form.TanggalPinjam)).Hours()/24)),
				form.FotoPath, pdf, doc, "",
			}

			vr := &sheets.ValueRange{Values: [][]interface{}{values}}
			_, err = sheetsService.Spreadsheets.Values.Update(sheetId, writeRange, vr).ValueInputOption("USER_ENTERED").Do()
			if err != nil {
				log.Println("❌ Gagal update data ke Sheets:", err)
				return
			}

			// Kirim WA
			salam := getSalam()
			pesan := fmt.Sprintf(`%s *%s* 👋

Terima kasih telah mengajukan izin pinjam alat dengan detail berikut:

🛠️ *Nama Alat*   : _%s_
📦 *Jumlah Alat* : _%d_
📅 *Tgl Pinjam*  : _%s_
📆 *Tgl Kembali* : _%s_

📄 *Berikut adalah dokumen peminjaman alat*: %s

⏳ Mohon tunggu persetujuan. Izin akan dikirim melalui WA ini.

🙏 Terima kasih.`, salam, form.Nama, form.NamaAlat, form.JumlahAlat, form.TanggalPinjam, form.TanggalKembali, pdf)

			log.Printf("DEBUG: Nomor WA yang akan dikirimi pesan (sebelum normalisasi): '%s'\n", form.NoWA)
			if form.NoWA == "" {
				log.Println("⚠️ Nomor WA peminjam kosong, tidak dapat mengirim pesan WA")
			} else {
				normalizedNo := normalizePhoneNumber(form.NoWA)
				log.Printf("DEBUG: Nomor WA setelah normalisasi: '%s'\n", normalizedNo)
				if normalizedNo == "" || !strings.HasPrefix(normalizedNo, "62") {
					log.Println("⚠️ Nomor WA peminjam tidak valid setelah normalisasi, tidak mengirim pesan WA")
				} else {
					err = kirimPesanWaBangkit(normalizedNo, pesan)
					if err != nil {
						log.Println("⚠️ Gagal kirim WA:", err)
					} else {
						log.Println("📲 WA terkirim ke:", normalizedNo)
					}
				}
			}

			// Kirim WA ke approver (nomor dan link approval diambil dari env atau config)
			approverNo := os.Getenv("APPROVER_NO")
			if approverNo == "" {
				approverNo = "6287760573989" // Default nomor approver jika env tidak ada
			}
			approvalLink := os.Getenv("APPROVAL_LINK")
			if approvalLink == "" {
				approvalLink = "https://example.com/approval" // Default link approval jika env tidak ada
			}
			approverPesan := fmt.Sprintf(`%s Bapak %s

%s telah mengajukan alat sebagai berikut : 
🛠️Nama Alat	:%s
📦Jml Alat	: %d	
📅Tgl pinjam   : %s
📅Tgl kembali  : %s

📄Berikut adalah dokumen peminjaman alat: %s

Mohon dapat memberikan persetujuan peminjaman alat melalui link berikut:
%s

🆔Untuk isian ID Peminjaman, silakan masukkan: %04d ✅

Terima kasih 🙏
`, salam, form.Nama, form.Nama, form.NamaAlat, form.JumlahAlat, form.TanggalPinjam, form.TanggalKembali, pdf, approvalLink, row)

			log.Printf("DEBUG: Mengirim WA ke approver dengan nomor: %s", approverNo)
			log.Printf("DEBUG: Pesan ke approver: %s", approverPesan)
			err = kirimPesanWaBangkit(approverNo, approverPesan)
			if err != nil {
				log.Printf("⚠️ Gagal kirim WA ke approver (%s): %v\n", approverNo, err)
			} else {
				log.Printf("📲 WA terkirim ke approver: %s\n", approverNo)
			}

			// Additional debug to confirm both messages sent
			log.Println("DEBUG: Selesai mengirim kedua pesan WA (peminjam dan approver)")
		}(form, localPath, sheetData)
}

func handleApprove(w http.ResponseWriter, r *http.Request) {
	r.ParseForm()
	idPinjam := r.FormValue("idPinjam")
	approver := r.FormValue("approver")
	statusPersetujuan := r.FormValue("statusPersetujuan")

	if idPinjam == "" || approver == "" || statusPersetujuan == "" {
		http.Error(w, "ID Pinjam, Approver, dan Status Persetujuan harus diisi", http.StatusBadRequest)
		return
	}

	sheetsService, _, _, err := getServices()
	if err != nil {
		http.Error(w, "Gagal inisialisasi layanan", http.StatusInternalServerError)
		log.Println("Service error:", err)
		return
	}

	sheetId := "1uULs6gLCAeLVeOI-qjdIcb4pRod-mC6g4Cu9TvtIVak"
	// Find the row with the matching idPinjam in column A (assuming idPinjam stored there)
	resp, err := sheetsService.Spreadsheets.Values.Get(sheetId, "Form Peminjam!A5:A").Do()
	if err != nil {
		http.Error(w, "Gagal mengambil data dari Sheets", http.StatusInternalServerError)
		log.Println("Sheets get error:", err)
		return
	}
	if resp == nil || resp.Values == nil {
		http.Error(w, "Data peminjaman kosong", http.StatusInternalServerError)
		log.Println("Empty peminjaman data")
		return
	}

	rowIndex := -1
	for i, row := range resp.Values {
		if len(row) > 0 && fmt.Sprintf("%v", row[0]) == idPinjam {
			rowIndex = i + 5 // Because range starts at row 5
			break
		}
	}
	if rowIndex == -1 {
		http.Error(w, "ID Pinjam tidak ditemukan", http.StatusBadRequest)
		return
	}

	// Update approval status in column Q (17th column, index 16)
	writeRangeStatus := fmt.Sprintf("Form Peminjam!Q%d", rowIndex)
	valuesStatus := [][]interface{}{{statusPersetujuan}}
	vrStatus := &sheets.ValueRange{Values: valuesStatus}
	_, err = sheetsService.Spreadsheets.Values.Update(sheetId, writeRangeStatus, vrStatus).ValueInputOption("USER_ENTERED").Do()
	if err != nil {
		http.Error(w, "Gagal update status approval", http.StatusInternalServerError)
		log.Println("Sheets update error (status):", err)
		return
	}

	// Update approval date in column R (18th column, index 17)
	writeRangeDate := fmt.Sprintf("Form Peminjam!R%d", rowIndex)
	valuesDate := [][]interface{}{{time.Now().Format("2006-01-02 15:04:05")}}
	vrDate := &sheets.ValueRange{Values: valuesDate}
	_, err = sheetsService.Spreadsheets.Values.Update(sheetId, writeRangeDate, vrDate).ValueInputOption("USER_ENTERED").Do()
	if err != nil {
		http.Error(w, "Gagal update tanggal persetujuan", http.StatusInternalServerError)
		log.Println("Sheets update error (date):", err)
		return
	}

	// Update approver name in column S (19th column, index 18)
	writeRangeApprover := fmt.Sprintf("Form Peminjam!S%d", rowIndex)
	valuesApprover := [][]interface{}{{approver}}
	vrApprover := &sheets.ValueRange{Values: valuesApprover}
	_, err = sheetsService.Spreadsheets.Values.Update(sheetId, writeRangeApprover, vrApprover).ValueInputOption("USER_ENTERED").Do()
	if err != nil {
		http.Error(w, "Gagal update nama approver", http.StatusInternalServerError)
		log.Println("Sheets update error (approver):", err)
		return
	}

	w.Write([]byte("✅ Approval berhasil dikirim"))
}

func generateSuratApproval(form FormData, nomorUrut int, approver, statusPersetujuan string, driveService *drive.Service, docsService *docs.Service) (pdfURL string, docURL string, err error) {
	templateID := "1NVr2LHlDrrqEJJTrCJed3AnQTncs5ZMU6Lu0wO1RlRs"
	pdfFolder := "1HhZncgqeqEzgTkMQZOBC9HAsPTIB0zTv"
	title := fmt.Sprintf("Formulir Approval %04d - %s", nomorUrut, form.Nama)

	// Salin template ke dokumen baru
	copy, err := driveService.Files.Copy(templateID, &drive.File{Name: title}).Do()
	if err != nil {
		log.Printf("❌ Gagal menyalin template: %v", err)
		return "", "", err
	}
	docID := copy.Id
	docURL = fmt.Sprintf("https://docs.google.com/document/d/%s/edit", docID)

	docFolder := "1Y3cvxCOy4M0GtRPe7A1DrAg1iji5O0lQ"
	_, err = driveService.Files.Update(docID, nil).
		AddParents(docFolder).
		RemoveParents("root").
		Do()
	if err != nil {
		log.Println("⚠️ Gagal memindahkan file ke folder Dokumen:", err)
	}

	// Siapkan teks pengganti
	replacements := map[string]string{
		"<<NMR>>":    fmt.Sprintf("%04d", nomorUrut),
		"<<TGL>>":    time.Now().Format("02 January 2006"),
		"<<NAMA>>":   form.Nama,
		"<<KLS>>":    form.Kelas,
		"<<NIS>>":    form.NIS,
		"<<NO>>":     form.NoWA,
		"<<NMALT>>":  form.NamaAlat,
		"<<JML>>":    strconv.Itoa(form.JumlahAlat),
		"<<TGLPMJ>>": form.TanggalPinjam,
		"<<TGLPGN>>": form.TanggalKembali,
		"<<LMPJM>>":  fmt.Sprintf("%d hari", int(parseTanggal(form.TanggalKembali).Sub(parseTanggal(form.TanggalPinjam)).Hours()/24)),
		"<<KET>>":    form.Keterangan,
		"<<TGLPS>>":  time.Now().Format("02 January 2006 15:04"),
		"<<STS>>":    statusPersetujuan,
		"<<YNG>>":    approver,
	}

	// Replace semua placeholder dalam dokumen
	var reqs []*docs.Request
	for key, val := range replacements {
		reqs = append(reqs, &docs.Request{
			ReplaceAllText: &docs.ReplaceAllTextRequest{
				ContainsText: &docs.SubstringMatchCriteria{Text: key, MatchCase: true},
				ReplaceText:  val,
			},
		})
	}
	_, err = docsService.Documents.BatchUpdate(docID, &docs.BatchUpdateDocumentRequest{Requests: reqs}).Do()
	if err != nil {
		log.Printf("❌ Gagal mengganti isi dokumen: %v", err)
		// Additional debug info
		log.Printf("DEBUG: BatchUpdate request: %+v", reqs)
		return "", "", err
	}

	// Replace <<FOTO>> placeholder with image if PeminjamanFotoPath is provided
	if form.PeminjamanFotoPath != "" {
		log.Printf("DEBUG: PeminjamanFotoPath is set: %s", form.PeminjamanFotoPath)
		doc, err := docsService.Documents.Get(docID).Do()
		if err != nil {
			log.Printf("ERROR: Failed to get document for <<FOTO>> replacement: %v", err)
		} else {
			var index int64 = -1
			for _, c := range doc.Body.Content {
				if c.Paragraph != nil {
					for _, e := range c.Paragraph.Elements {
						if e.TextRun != nil && strings.Contains(e.TextRun.Content, "<<FOTO>>") {
							index = e.StartIndex
							log.Printf("DEBUG: Found <<FOTO>> placeholder at index %d", index)
							break
						}
					}
				}
				if index != -1 {
					break
				}
			}
			if index == -1 {
				log.Println("WARNING: <<FOTO>> placeholder not found in document")
			} else {
				end := index + int64(len("<<FOTO>>"))
				imgReq := []*docs.Request{
					{DeleteContentRange: &docs.DeleteContentRangeRequest{
						Range: &docs.Range{StartIndex: index, EndIndex: end},
					}},
					{InsertInlineImage: &docs.InsertInlineImageRequest{
						Location: &docs.Location{Index: index},
						Uri:      form.PeminjamanFotoPath,
					ObjectSize: &docs.Size{
						Width:  &docs.Dimension{Magnitude: 400, Unit: "PT"},
						Height: &docs.Dimension{Magnitude: 225, Unit: "PT"},
					},
					}},
				}
				resp, err := docsService.Documents.BatchUpdate(docID, &docs.BatchUpdateDocumentRequest{Requests: imgReq}).Do()
				if err != nil {
					log.Printf("ERROR: Failed to insert image for <<FOTO>> placeholder: %v", err)
				} else {
					log.Printf("DEBUG: Successfully inserted image for <<FOTO>> placeholder, response: %+v", resp)
				}
			}
		}
	} else {
		log.Println("DEBUG: PeminjamanFotoPath is empty, skipping <<FOTO>> replacement")
	}

	// Jadikan dokumen publik
	_, _ = driveService.Permissions.Create(docID, &drive.Permission{
		Type: "anyone",
		Role: "reader",
	}).Do()

	// Export to PDF
	export, err := driveService.Files.Export(docID, "application/pdf").Download()
	if err != nil {
		log.Printf("❌ Gagal export PDF: %v", err)
		return "", "", err
	}
	tmp := filepath.Join("uploads", fmt.Sprintf("%s.pdf", title))
	out, _ := os.Create(tmp)
	io.Copy(out, export.Body)
	out.Close()

	file, _ := os.Open(tmp)
	pdf, err := driveService.Files.Create(&drive.File{
		Name:     filepath.Base(tmp),
		Parents:  []string{pdfFolder},
		MimeType: "application/pdf",
	}).Media(file).Do()
	file.Close()
	os.Remove(tmp)

	if err != nil {
		log.Printf("❌ Gagal upload PDF: %v", err)
		return "", "", err
	}

	driveService.Permissions.Create(pdf.Id, &drive.Permission{Role: "reader", Type: "anyone"}).Do()

	pdfURL = fmt.Sprintf("https://drive.google.com/uc?id=%s", pdf.Id)

	log.Printf("✅ Dokumen approval berhasil dibuat: %s", docURL)
	log.Printf("✅ PDF approval berhasil dibuat: %s", pdfURL)
	return pdfURL, docURL, nil
}


func handleApprovalRequestNew(w http.ResponseWriter, r *http.Request) {
	r.ParseMultipartForm(10 << 20)
	idPinjam := r.FormValue("idPinjam")
	approver := r.FormValue("approver")
	statusPersetujuan := r.FormValue("statusPersetujuan")

	log.Printf("DEBUG: Received approval request with idPinjam: '%s', approver: '%s', statusPersetujuan: '%s'\n", idPinjam, approver, statusPersetujuan)

	if idPinjam == "" || approver == "" || statusPersetujuan == "" {
		http.Error(w, "ID Pinjam, Approver, dan Status Persetujuan harus diisi", http.StatusBadRequest)
		return
	}

	sheetsService, driveService, docsService, err := getServices()
	if err != nil {
		http.Error(w, "Gagal inisialisasi layanan", http.StatusInternalServerError)
		log.Println("Service error:", err)
		return
	}

	// Find the row with the matching idPinjam in the "Form Peminjam" sheet to get peminjam details
	sheetId := "1uULs6gLCAeLVeOI-qjdIcb4pRod-mC6g4Cu9TvtIVak"
	resp, err := sheetsService.Spreadsheets.Values.Get(sheetId, "Form Peminjam!A5:Z").Do()
	if err != nil {
		http.Error(w, "Gagal mengambil data dari Sheets", http.StatusInternalServerError)
		log.Println("Sheets get error:", err)
		return
	}
	if resp == nil || resp.Values == nil {
		http.Error(w, "Data peminjaman kosong", http.StatusInternalServerError)
		log.Println("Empty peminjaman data")
		return
	}

	var peminjamName, noWA, kelas, nis, namaAlat, tglPinjam, tglKembali, keterangan string
	var jumlahAlat int
	for _, row := range resp.Values {
		if len(row) > 0 {
			sheetID := fmt.Sprintf("%v", row[0])
			sheetIDTrimmed := strings.TrimLeft(sheetID, "0")
			idPinjamTrimmed := strings.TrimLeft(idPinjam, "0")
			if sheetIDTrimmed == idPinjamTrimmed {
				if len(row) > 2 {
					peminjamName = fmt.Sprintf("%v", row[2])
				}
				if len(row) > 3 {
					kelas = fmt.Sprintf("%v", row[3])
				}
				if len(row) > 4 {
					nis = fmt.Sprintf("%v", row[4])
				}
				if len(row) > 5 {
					noWA = strings.TrimSpace(fmt.Sprintf("%v", row[5]))
				}
				if len(row) > 6 {
					namaAlat = fmt.Sprintf("%v", row[6])
				}
				if len(row) > 7 {
					jumlahAlat, _ = strconv.Atoi(fmt.Sprintf("%v", row[7]))
				}
				if len(row) > 8 {
					tglPinjam = fmt.Sprintf("%v", row[8])
				}
				if len(row) > 9 {
					tglKembali = fmt.Sprintf("%v", row[9])
				}
				if len(row) > 11 {
					keterangan = fmt.Sprintf("%v", row[11])
				}
				break
			}
		}
	}

	// Calculate lamaPinjam
	layout := "2006-01-02"
	start, errStart := time.Parse(layout, tglPinjam)
	end, errEnd := time.Parse(layout, tglKembali)
	if errStart == nil && errEnd == nil {
		days := int(end.Sub(start).Hours() / 24)
		if days < 0 {
			days = 0
		}
		// lamaPinjam variable removed as it was unused
	} 

	// Prepare form data for document generation
	form := FormData{
		Nama:               peminjamName,
		Kelas:              kelas,
		NIS:                nis,
		NoWA:               noWA,
		NamaAlat:           namaAlat,
		JumlahAlat:         jumlahAlat,
		TanggalPinjam:      tglPinjam,
		TanggalKembali:     tglKembali,
		Keterangan:         keterangan,
		PeminjamanFotoPath: "", // Initialize empty
	}

	// Attempt to set PeminjamanFotoPath from the sheet data row
	for _, row := range resp.Values {
		if len(row) > 0 {
			sheetID := fmt.Sprintf("%v", row[0])
			sheetIDTrimmed := strings.TrimLeft(sheetID, "0")
			idPinjamTrimmed := strings.TrimLeft(idPinjam, "0")
			if sheetIDTrimmed == idPinjamTrimmed {
				if len(row) > 12 {
					form.PeminjamanFotoPath = fmt.Sprintf("%v", row[12])
					log.Printf("DEBUG: PeminjamanFotoPath set from sheet: %s", form.PeminjamanFotoPath)
				}
				break
			}
		}
	}

	// Convert idPinjam to int nomorUrut for approval document
	nomorUrut, convErr := strconv.Atoi(idPinjam)
	if convErr != nil {
		log.Printf("⚠️ Gagal konversi idPinjam ke int: %v, menggunakan 1 sebagai default", convErr)
		nomorUrut = 1
	}

	// Generate approval document using the existing function with updated templateID
	docURL, _, err := generateSuratApproval(form, nomorUrut, approver, statusPersetujuan, driveService, docsService)
	if err != nil {
		http.Error(w, "Gagal membuat dokumen approval", http.StatusInternalServerError)
		log.Println("generateSuratApproval error:", err)
		return
	}

	// Insert data into "Approval Peminjaman" sheet, tab "Approval Peminjaman"
	approvalSheetId := "1uULs6gLCAeLVeOI-qjdIcb4pRod-mC6g4Cu9TvtIVak"
	approvalSheetRange := "Approval Peminjaman!A6:F"
	respApproval, err := sheetsService.Spreadsheets.Values.Get(approvalSheetId, approvalSheetRange).Do()
	if err != nil {
		http.Error(w, "Gagal mengambil data dari sheet approval", http.StatusInternalServerError)
		log.Println("Sheets get error:", err)
		return
	}

	rowNum := 6
	if respApproval != nil && respApproval.Values != nil {
		rowNum = len(respApproval.Values) + 6
	}

	writeRange := fmt.Sprintf("Approval Peminjaman!A%d", rowNum)

	values := []interface{}{
		fmt.Sprintf("%04d", rowNum-5),
		time.Now().Format("2006-01-02"),
		peminjamName,
		approver,
		idPinjam,
		statusPersetujuan,
	}
	vr := &sheets.ValueRange{Values: [][]interface{}{values}}
	_, err = sheetsService.Spreadsheets.Values.Update(approvalSheetId, writeRange, vr).ValueInputOption("USER_ENTERED").Do()
	if err != nil {
		http.Error(w, "Gagal update data ke sheet approval", http.StatusInternalServerError)
		log.Println("Sheets update error:", err)
		return
	}

	// Send WhatsApp notifications to peminjam and approver
	// For peminjam, strictly get NoWA from "Form Peminjam" sheet column F (index 5)
	var noWAApproval string
	noWASet := false
	idPinjamTrimmed := strings.TrimLeft(idPinjam, "0")

	for _, row := range resp.Values {
		if len(row) > 0 {
			sheetID := fmt.Sprintf("%v", row[0])
			sheetIDTrimmed := strings.TrimLeft(sheetID, "0")

			if sheetIDTrimmed == idPinjamTrimmed {
				if len(row) > 5 {
					noWAApproval = strings.TrimSpace(fmt.Sprintf("%v", row[5]))
					noWASet = true
				}
				break
			}
		}
	}

	if !noWASet {
		noWAApproval = ""
	}

	salam := getSalam()
	pesanPeminjam := fmt.Sprintf(`%s %s

Pengajuan peminjaman alat berikut:

Nama Alat       : %s
Jumlah Alat     : %d
Tgl Pinjam      : %s
Tgl Harus Kembali : %s
Status Persetujuan : %s
Pemberi ijin    : Bapak %s

Silahkan gunakan alat dengan baik.
Jika sudah selesai digunakan silahkan isi formulir pengembalian alat melalui link berikut: https://s.id/FormKembaliAlat

Dokumen persetujuan:
%s

Terima Kasih 🙏`, salam, peminjamName, namaAlat, jumlahAlat, tglPinjam, tglKembali, statusPersetujuan, approver, docURL)

	normalizedNoWA := normalizePhoneNumber(noWAApproval)
	if normalizedNoWA == "" || !strings.HasPrefix(normalizedNoWA, "62") {
		log.Println("⚠️ Nomor WA peminjam untuk approval tidak valid, tidak mengirim pesan WA")
		// Fallback: try to get NoWA from form peminjam sheet by matching idPinjam again
		noWAFallback := ""
		for _, row := range resp.Values {
			if len(row) > 0 && fmt.Sprintf("%v", row[0]) == idPinjam {
				if len(row) > 5 {
					noWAFallback = strings.TrimSpace(fmt.Sprintf("%v", row[5]))
				}
				break
			}
		}
		normalizedFallback := normalizePhoneNumber(noWAFallback)
		if normalizedFallback != "" && strings.HasPrefix(normalizedFallback, "62") {
			err = kirimPesanWaBangkit(normalizedFallback, pesanPeminjam)
			if err != nil {
				log.Println("⚠️ Gagal kirim WA ke peminjam dengan fallback:", err)
			} else {
				log.Println("📲 WA terkirim ke peminjam dengan fallback:", normalizedFallback)
			}
		}
	} else {
		err = kirimPesanWaBangkit(normalizedNoWA, pesanPeminjam)
		if err != nil {
			log.Println("⚠️ Gagal kirim WA ke peminjam:", err)
		} else {
			log.Println("📲 WA terkirim ke peminjam:", normalizedNoWA)
		}
	}

	// Send WA to approver
	approverNo := os.Getenv("APPROVER_NO")
	if approverNo == "" {
		approverNo = "6287760573989"
	}
	pesanApprover := fmt.Sprintf(`%s Bapak/Ibu %s

Permohonan persetujuan dengan ID %s dari %s telah diproses dengan status: %s.

📄 Dokumen persetujuan: %s

Terima kasih.`, salam, approver, idPinjam, peminjamName, statusPersetujuan, docURL)

	err = kirimPesanWaBangkit(approverNo, pesanApprover)
	if err != nil {
		log.Println("⚠️ Gagal kirim WA ke approver:", err)
	} else {
		log.Println("📲 WA terkirim ke approver:", approverNo)
	}

	w.Write([]byte("✅ Permohonan persetujuan berhasil diproses"))
}



func handlePengembalian(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Access-Control-Allow-Origin", "*")

	if r.Method == "OPTIONS" {
		w.WriteHeader(http.StatusOK)
		return
	}

	if r.Method != "POST" {
		http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
		return
	}

	r.ParseMultipartForm(10 << 20)

	idPeminjam := r.FormValue("idPeminjam")
	kondisiAlat := r.FormValue("kondisiAlat")
	keteranganPengembalian := r.FormValue("keteranganPengembalian")

	// Save the uploaded file locally first
	var localPath string
	file, handler, err := r.FormFile("foto")
	if err == nil {
		defer file.Close()
		localPath, _ = saveFileLocally(file, handler.Filename)
	}

	// Respond immediately to the client
	w.Write([]byte("✅ Data pengembalian berhasil diterima dan sedang diproses"))

	go func(idPeminjam, kondisiAlat, keteranganPengembalian, localPath string) {
		sheetsService, driveService, docsService, err := getServices()
		if err != nil {
			log.Println("Service error:", err)
			return
		}

	// Fetch peminjaman details by idPeminjam from "Form Peminjam" sheet
	sheetId := "1uULs6gLCAeLVeOI-qjdIcb4pRod-mC6g4Cu9TvtIVak"
	resp, err := sheetsService.Spreadsheets.Values.Get(sheetId, "Form Peminjam!A5:Z").Do()
	if err != nil {
		log.Println("❌ Gagal mengambil data dari Sheets:", err)
		return
	}

	var form FormData
	found := false
	for _, row := range resp.Values {
		if len(row) > 0 {
			sheetID := fmt.Sprintf("%v", row[0])
			sheetIDTrimmed := strings.TrimLeft(sheetID, "0")
			idPeminjamTrimmed := strings.TrimLeft(idPeminjam, "0")
			if sheetIDTrimmed == idPeminjamTrimmed {
				form.Nama = fmt.Sprintf("%v", row[2])
				form.Kelas = fmt.Sprintf("%v", row[3])
				form.NIS = fmt.Sprintf("%v", row[4])
				form.NoWA = fmt.Sprintf("%v", row[5])
				form.NamaAlat = fmt.Sprintf("%v", row[6])
				form.JumlahAlat, _ = strconv.Atoi(fmt.Sprintf("%v", row[7]))
				form.TanggalPinjam = fmt.Sprintf("%v", row[8])
				form.TanggalKembali = fmt.Sprintf("%v", row[9])
				form.KondisiAlat = kondisiAlat
				form.KeteranganPengembalian = keteranganPengembalian
				if len(row) > 10 {
					form.KeteranganPinjam = fmt.Sprintf("%v", row[10])
					log.Printf("DEBUG: KeteranganPinjam read from sheet: '%s'", form.KeteranganPinjam)
				}
				if len(row) > 12 {
					form.PeminjamanFotoPath = fmt.Sprintf("%v", row[12])
					log.Printf("DEBUG: PeminjamanFotoPath read from sheet: '%s'", form.PeminjamanFotoPath)
				}

	// Fetch approval data from "Approval Peminjaman" sheet
	approvalSheetId := sheetId
	approvalRange := "Approval Peminjaman!A6:F"
	respApproval, err := sheetsService.Spreadsheets.Values.Get(approvalSheetId, approvalRange).Do()
	if err != nil {
		log.Println("❌ Gagal mengambil data dari sheet Approval Peminjaman:", err)
	} else if respApproval != nil && respApproval.Values != nil {
		idPeminjamTrimmed := strings.TrimLeft(idPeminjam, "0")
		for _, approvalRow := range respApproval.Values {
			if len(approvalRow) > 4 {
				approvalId := fmt.Sprintf("%v", approvalRow[4])
				approvalIdTrimmed := strings.TrimLeft(approvalId, "0")
				if approvalIdTrimmed == idPeminjamTrimmed {
					if len(approvalRow) > 1 {
						form.ApprovalDate = fmt.Sprintf("%v", approvalRow[1])
					}
					if len(approvalRow) > 5 {
						form.ApprovalStatus = fmt.Sprintf("%v", approvalRow[5])
					}
					if len(approvalRow) > 3 {
						form.ApproverName = fmt.Sprintf("%v", approvalRow[3])
					}
					break
				}
			}
		}
	}

					found = true
					break
				}
			}
		}

		if !found {
			log.Println("❌ ID Peminjam tidak ditemukan di sheet peminjaman")
			return
		}

		// Upload file to Drive if available
		if localPath != "" {
			url, err := uploadToDrive(localPath, filepath.Base(localPath), driveService)
			if err == nil {
				form.FotoPath = url
				log.Println("✅ Foto pengembalian berhasil diupload:", form.FotoPath)
			} else {
				log.Println("❌ Gagal upload foto pengembalian ke Drive:", err)
				form.FotoPath = "Gagal upload"
			}
			os.Remove(localPath)
		}

		// Use the same sheet ID but different sheet name "Form Pengembalian"
		respPengembalian, err := sheetsService.Spreadsheets.Values.Get(sheetId, "Form Pengembalian!B5:B").Do()
		if err != nil {
			log.Println("❌ Gagal mengambil data dari Sheets pengembalian:", err)
			return
		}

		var row int
		if respPengembalian == nil || respPengembalian.Values == nil || len(respPengembalian.Values) == 0 {
			log.Println("❌ Response dari Sheets pengembalian kosong, memulai dari baris 1")
			row = 1
		} else {
			row = len(respPengembalian.Values) + 1
		}

		writeRange := fmt.Sprintf("Form Pengembalian!A%d", row+4)

		// Convert idPeminjam to int for nomorUrut
		nomorUrut := row
		idPeminjamInt, errConv := strconv.Atoi(idPeminjam)
		if errConv == nil {
			nomorUrut = idPeminjamInt
		}

		// Generate surat pengembalian using the correct function
		pdf, _, err := generateSuratPengembalian(form, nomorUrut, driveService, docsService)
		if err != nil {
			log.Println("❌ Gagal generate surat pengembalian:", err)
			return
		}

		// Convert idPeminjam to int for consistent formatting
		idPeminjamInt, errConv = strconv.Atoi(idPeminjam)
		idPeminjamFormatted := idPeminjam
		if errConv == nil {
			idPeminjamFormatted = fmt.Sprintf("%04d", idPeminjamInt)
		}

		values := []interface{}{
			idPeminjamFormatted,           // Kolom A: ID PEMINJAM
			form.Nama,                     // Kolom B: NAMA
			time.Now().Format("2006-01-02"), // Kolom C: TANGGAL PENGEMBALIAN
			kondisiAlat,                   // Kolom D: KONDISI ALAT
			keteranganPengembalian,        // Kolom E: KETERANGAN
			form.FotoPath,                 // Kolom F: UP FOTO PENGEMBALIAN
		}

		log.Printf("DEBUG: ID: %s | Nama: %s | Kondisi: %s | Ket: %s", idPeminjam, form.Nama, kondisiAlat, keteranganPengembalian)
		log.Printf("DEBUG: Writing to Form Pengembalian sheet at range %s with values: %+v", writeRange, values)

		vr := &sheets.ValueRange{Values: [][]interface{}{values}}
		respUpdate, err := sheetsService.Spreadsheets.Values.Update(sheetId, writeRange, vr).ValueInputOption("USER_ENTERED").Do()
		if err != nil {
			log.Println("❌ Gagal update data pengembalian ke Sheets:", err)
			return
		} else {
			log.Printf("INFO: Update response from Sheets API: %+v", respUpdate)
		}

		// Kirim WA notifikasi ke peminjam
		salam := getSalam()
		pesan := fmt.Sprintf(`%s *%s* 👋

Terima kasih telah melakukan pengembalian alat dengan detail berikut:

🛠️ *Nama Alat*   : _%s_
📦 *Jumlah Alat* : _%d_
📅 *Tgl Pinjam*  : _%s_
📆 *Tgl Kembali* : _%s_
📋 *Kondisi Alat*: _%s_

📄 *Dokumen Pengembalian*: %s

🙏 Terima kasih.`, salam, form.Nama, form.NamaAlat, form.JumlahAlat, form.TanggalPinjam, form.TanggalKembali, kondisiAlat, pdf)

		if form.NoWA == "" {
			log.Println("⚠️ Nomor WA peminjam kosong, tidak dapat mengirim pesan WA")
		} else {
			normalizedNo := normalizePhoneNumber(form.NoWA)
			if normalizedNo == "" || !strings.HasPrefix(normalizedNo, "62") {
				log.Println("⚠️ Nomor WA peminjam tidak valid setelah normalisasi, tidak mengirim pesan WA")
			} else {
				err = kirimPesanWaBangkit(normalizedNo, pesan)
				if err != nil {
					log.Println("⚠️ Gagal kirim WA:", err)
				} else {
					log.Println("📲 WA pengembalian terkirim ke:", normalizedNo)
				}
			}
		}

		// Kirim WA notifikasi ke approver
		approverNo := os.Getenv("APPROVER_NO")
		if approverNo == "" {
			approverNo = "6287760573989" // Default nomor approver jika env tidak ada
		}

		// Use approver name from approval sheet if available, else fallback to "Bapak Sebastian"
		approverName := "Bapak Sebastian"
		if form.ApproverName != "" {
			approverName = form.ApproverName
		}

		// Use current date as Tgl Kembali in message
		tglKembaliNow := time.Now().Format("02 January 2006")

		pesanApprover := fmt.Sprintf(`Selamat Malam %s

Melaporkan, %s telah mengembalikan alat berikut:

Nama Alat       : %s
Jumlah Alat     : %d
Tgl Pinjam       : %s
Tgl Harus Kembali   : %s
Tgl Kembali     : %s
Kondisi Alat     : %s
Keterangan      : %s

Berikut dokumen pengembalian alat:
%s

Terima Kasih 🙏
`, approverName, form.Nama, form.NamaAlat, form.JumlahAlat, form.TanggalPinjam, form.TanggalKembali, tglKembaliNow, kondisiAlat, keteranganPengembalian, pdf)

		normalizedApproverNo := normalizePhoneNumber(approverNo)
		if normalizedApproverNo == "" || !strings.HasPrefix(normalizedApproverNo, "62") {
			log.Println("⚠️ Nomor WA approver tidak valid, tidak mengirim pesan WA")
		} else {
			err = kirimPesanWaBangkit(normalizedApproverNo, pesanApprover)
			if err != nil {
				log.Println("⚠️ Gagal kirim WA ke approver:", err)
			} else {
				log.Println("📲 WA pengembalian terkirim ke approver:", normalizedApproverNo)
			}
		}

	}(idPeminjam, kondisiAlat, keteranganPengembalian, localPath)
}

func generateSuratPengembalian(form FormData, nomorUrut int, driveService *drive.Service, docsService *docs.Service) (pdfURL, docURL string, err error) {
	templateID := "1aBpU0yBFFjVdMjYtuB5skHY4m5pCKlVlMCdzq5Ib9Y0"
	pdfFolder := "1HhZncgqeqEzgTkMQZOBC9HAsPTIB0zTv"
	title := fmt.Sprintf("Formulir Pengembalian %04d - %s", nomorUrut, form.Nama)

	copy, err := driveService.Files.Copy(templateID, &drive.File{Name: title}).Do()
	if err != nil {
		return "", "", fmt.Errorf("❌ Gagal menyalin template: %v", err)
	}
	docID := copy.Id
	docURL = fmt.Sprintf("https://docs.google.com/document/d/%s/edit", docID)

	docFolder := "1Y3cvxCOy4M0GtRPe7A1DrAg1iji5O0lQ"
	_, _ = driveService.Files.Update(docID, nil).
		AddParents(docFolder).
		RemoveParents("root").
		Do()

	replacements := map[string]string{
		"<<NMR>>":     fmt.Sprintf("%04d", nomorUrut),
		"<<TGL>>":     time.Now().Format("02 January 2006"),
		"<<TGLBALI>>": time.Now().Format("02 January 2006"),
		"<<NAMA>>":    form.Nama,
		"<<KLS>>":     form.Kelas,
		"<<NIS>>":     form.NIS,
		"<<NO>>":      form.NoWA,
		"<<NMALT>>":   form.NamaAlat,
		"<<JML>>":     strconv.Itoa(form.JumlahAlat),
		"<<TGLPMJ>>":  form.TanggalPinjam,
		"<<TGLPGN>>":  form.TanggalKembali,
		"<<LMPJM>>":   fmt.Sprintf("%d hari", int(parseTanggal(form.TanggalKembali).Sub(parseTanggal(form.TanggalPinjam)).Hours()/24)),
		"<<KET>>":     form.KeteranganPinjam,
		"<<KNDS>>":    form.KondisiAlat,
		"<<KETALT>>":  form.KeteranganPengembalian,
		"<<TGLPS>>":   form.ApprovalDate,
		"<<STS>>":     form.ApprovalStatus,
		"<<YNG>>":     form.ApproverName,
	}

	var reqs []*docs.Request
	for key, val := range replacements {
		reqs = append(reqs, &docs.Request{
			ReplaceAllText: &docs.ReplaceAllTextRequest{
				ContainsText: &docs.SubstringMatchCriteria{Text: key, MatchCase: true},
				ReplaceText:  val,
			},
		})
	}
	_, err = docsService.Documents.BatchUpdate(docID, &docs.BatchUpdateDocumentRequest{Requests: reqs}).Do()
	if err != nil {
		return "", "", fmt.Errorf("❌ Gagal mengganti isi dokumen: %v", err)
	}

	// Tambahkan foto jika tersedia
	if form.PeminjamanFotoPath != "" {
		doc, err := docsService.Documents.Get(docID).Do()
		if err == nil {
			var index int64
			for _, c := range doc.Body.Content {
				if c.Paragraph != nil {
					for _, e := range c.Paragraph.Elements {
						if e.TextRun != nil && strings.Contains(e.TextRun.Content, "<<FOTO>>") {
							index = e.StartIndex
							break
						}
					}
				}
			}
			end := index + int64(len("<<FOTO>>"))
			imgReq := []*docs.Request{
				{DeleteContentRange: &docs.DeleteContentRangeRequest{
					Range: &docs.Range{StartIndex: index, EndIndex: end},
				}},
				{InsertInlineImage: &docs.InsertInlineImageRequest{
					Location: &docs.Location{Index: index},
					Uri:      form.PeminjamanFotoPath,
					ObjectSize: &docs.Size{
						Width:  &docs.Dimension{Magnitude: 400, Unit: "PT"},
						Height: &docs.Dimension{Magnitude: 225, Unit: "PT"},
					},
				}},
			}
			docsService.Documents.BatchUpdate(docID, &docs.BatchUpdateDocumentRequest{Requests: imgReq}).Do()
		}

		// Tambahkan foto kedua <<FOTO2>> jika tersedia
		var index2 int64
		foundFoto2 := false
		for _, c := range doc.Body.Content {
			if c.Paragraph != nil {
				for _, e := range c.Paragraph.Elements {
					if e.TextRun != nil && strings.Contains(e.TextRun.Content, "<<FOTO2>>") {
						index2 = e.StartIndex
						foundFoto2 = true
						log.Printf("DEBUG: Found <<FOTO2>> placeholder at index %d", index2)
						break
					}
				}
			}
		}
		if !foundFoto2 {
			log.Println("DEBUG: <<FOTO2>> placeholder not found in document")
		}
		if form.FotoPath == "" {
			log.Println("DEBUG: form.FotoPath is empty, cannot replace <<FOTO2>>")
		} else {
			log.Printf("DEBUG: Replacing <<FOTO2>> with image URL: %s", form.FotoPath)
			// Replace <<FOTO2>> with a unique marker text "IMG_PLACEHOLDER"
			replaceReq := []*docs.Request{
				{
					ReplaceAllText: &docs.ReplaceAllTextRequest{
						ContainsText: &docs.SubstringMatchCriteria{
							Text:      "<<FOTO2>>",
							MatchCase: true,
						},
						ReplaceText: "IMG_PLACEHOLDER",
					},
				},
			}
			_, err := docsService.Documents.BatchUpdate(docID, &docs.BatchUpdateDocumentRequest{Requests: replaceReq}).Do()
			if err != nil {
				log.Printf("ERROR: Failed to replace <<FOTO2>> placeholder with marker: %v", err)
				return "", "", err
			}
			// Fetch document content again to find index of "IMG_PLACEHOLDER"
			doc, err := docsService.Documents.Get(docID).Do()
			if err != nil {
				log.Printf("ERROR: Failed to fetch document after replacing <<FOTO2>>: %v", err)
				return "", "", err
			}
			var markerIndex int64 = -1
			for _, c := range doc.Body.Content {
				if c.Paragraph != nil {
					for _, e := range c.Paragraph.Elements {
						if e.TextRun != nil && strings.Contains(e.TextRun.Content, "IMG_PLACEHOLDER") {
							markerIndex = e.StartIndex
							break
						}
					}
				}
				if markerIndex != -1 {
					break
				}
			}
			if markerIndex == -1 {
				log.Printf("ERROR: Marker 'IMG_PLACEHOLDER' not found in document")
				return "", "", fmt.Errorf("marker 'IMG_PLACEHOLDER' not found")
			}
			// Batch update to delete marker and insert image at markerIndex
			imgReq2 := []*docs.Request{
				{
					DeleteContentRange: &docs.DeleteContentRangeRequest{
						Range: &docs.Range{
							StartIndex: markerIndex,
							EndIndex:   markerIndex + int64(len("IMG_PLACEHOLDER")),
						},
					},
				},
				{
					InsertInlineImage: &docs.InsertInlineImageRequest{
						Location: &docs.Location{Index: markerIndex},
						Uri:      form.FotoPath, // This is the pengembalian photo URL
					ObjectSize: &docs.Size{
						Width:  &docs.Dimension{Magnitude: 400, Unit: "PT"},
						Height: &docs.Dimension{Magnitude: 225, Unit: "PT"},
					},
					},
				},
			}
			_, err = docsService.Documents.BatchUpdate(docID, &docs.BatchUpdateDocumentRequest{Requests: imgReq2}).Do()
			if err != nil {
				log.Printf("ERROR: Failed to insert image for <<FOTO2>> placeholder: %v", err)
			} else {
				log.Println("DEBUG: Successfully inserted image for <<FOTO2>> placeholder")
			}
		}
	}

	// Buat PDF dari dokumen
	export, err := driveService.Files.Export(docID, "application/pdf").Download()
	if err != nil {
		return "", "", fmt.Errorf("❌ Gagal export PDF: %v", err)
	}
	tmp := filepath.Join("uploads", fmt.Sprintf("%s.pdf", title))
	out, _ := os.Create(tmp)
	io.Copy(out, export.Body)
	out.Close()

	file, _ := os.Open(tmp)
	pdf, _ := driveService.Files.Create(&drive.File{
		Name:     filepath.Base(tmp),
		Parents:  []string{pdfFolder},
		MimeType: "application/pdf",
	}).Media(file).Do()
	file.Close()
	os.Remove(tmp)

	driveService.Permissions.Create(pdf.Id, &drive.Permission{Role: "reader", Type: "anyone"}).Do()
	driveService.Permissions.Create(docID, &drive.Permission{Role: "reader", Type: "anyone"}).Do()

	pdfURL = fmt.Sprintf("https://drive.google.com/uc?id=%s", pdf.Id)
	return pdfURL, docURL, nil
}

func main() {
	http.HandleFunc("/pinjam", handlePinjam)
	http.HandleFunc("/approve", handleApprove)
	http.HandleFunc("/approval-request-new", handleApprovalRequestNew)
	http.HandleFunc("/pengembalian", handlePengembalian)
	fmt.Println("🚀 Server berjalan di http://localhost:8080")
	log.Fatal(http.ListenAndServe(":8080", cors.AllowAll().Handler(http.DefaultServeMux)))
}
