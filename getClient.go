package main

import (
	"context"
	"encoding/json"
	"fmt"
	"log"
	"net/http"
	"os"
	"os/user"
	"path/filepath"

	"golang.org/x/oauth2"
)

func getClient(config *oauth2.Config) *http.Client {
	tokFile := tokenCacheFile()
	tok, err := tokenFromFile(tokFile)
	if err != nil {
		tok = getTokenFromWeb(config)
		saveToken(tokFile, tok)
	}
	return config.Client(context.Background(), tok)
}

func tokenCacheFile() string {
	usr, err := user.Current()
	if err != nil {
		log.Fatalf("❌ Gagal ambil user: %v", err)
	}
	tokenDir := filepath.Join(usr.HomeDir, ".credentials")
	os.MkdirAll(tokenDir, 0700)
	return filepath.Join(tokenDir, "token.json")
}

func tokenFromFile(file string) (*oauth2.Token, error) {
	f, err := os.Open(file)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	tok := &oauth2.Token{}
	err = json.NewDecoder(f).Decode(tok)
	return tok, err
}

func getTokenFromWeb(config *oauth2.Config) *oauth2.Token {
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	fmt.Printf("👉 Silakan buka link ini di browser dan login:\n%v\n", authURL)
	fmt.Print("📝 Masukkan kode otorisasi di sini: ")

	var authCode string
	if _, err := fmt.Scan(&authCode); err != nil {
		log.Fatalf("❌ Gagal membaca kode: %v", err)
	}

	tok, err := config.Exchange(context.TODO(), authCode)
	if err != nil {
		log.Fatalf("❌ Gagal menukar kode dengan token: %v", err)
	}
	return tok
}

func saveToken(path string, token *oauth2.Token) {
	fmt.Printf("✅ Token disimpan ke: %s\n", path)
	f, err := os.Create(path)
	if err != nil {
		log.Fatalf("❌ Gagal menyimpan token: %v", err)
	}
	defer f.Close()
	json.NewEncoder(f).Encode(token)
}
