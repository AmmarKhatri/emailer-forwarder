package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"net/http"
)

const (
	ClientID     = "b8ebdc42-2414-425d-95e2-d70b9f94fb43"
	ClientSecret = "ynC8Q~j1e9HkIjdpJXVfzHHhCZllaVxdQ~2J8bGi"
	TenantID     = "57bd375a-8f5a-4585-8cd4-9c82ba31f845"
	GraphAPI     = "https://graph.microsoft.com/v1.0"
	TokenURL     = "https://login.microsoftonline.com/%s/oauth2/v2.0/token"
)

var AccessToken string

type TokenResponse struct {
	AccessToken string `json:"access_token"`
}

type Email struct {
	ID      string `json:"id"`
	Subject string `json:"subject"`
	Body    struct {
		Content string `json:"content"`
	} `json:"body"`
	Sender struct {
		EmailAddress struct {
			Address string `json:"address"`
		} `json:"emailAddress"`
	} `json:"sender"`
}

func getAccessToken() error {
	url := fmt.Sprintf(TokenURL, TenantID)
	data := "client_id=" + ClientID + "&client_secret=" + ClientSecret + "&scope=https://graph.microsoft.com/.default&grant_type=client_credentials"
	resp, err := http.Post(url, "application/x-www-form-urlencoded", bytes.NewBufferString(data))
	if err != nil {
		return err
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return fmt.Errorf("failed to get token: %s", resp.Status)
	}

	var tokenResp TokenResponse
	if err := json.NewDecoder(resp.Body).Decode(&tokenResp); err != nil {
		return err
	}

	AccessToken = tokenResp.AccessToken
	return nil
}

func getFolderID(userID, parentFolderName, targetFolderName string) (string, error) {
	// Step 1: Get the parent folder ID (e.g., "Inbox")
	parentFolderURL := fmt.Sprintf("%s/users/%s/mailFolders", GraphAPI, userID)
	req, err := http.NewRequest("GET", parentFolderURL, nil)
	if err != nil {
		return "", err
	}
	req.Header.Set("Authorization", "Bearer "+AccessToken)

	resp, err := http.DefaultClient.Do(req)
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		body, _ := ioutil.ReadAll(resp.Body)
		return "", fmt.Errorf("failed to get mail folders: %s, %s", resp.Status, string(body))
	}

	var folders struct {
		Value []struct {
			ID   string `json:"id"`
			Name string `json:"displayName"`
		} `json:"value"`
	}
	if err := json.NewDecoder(resp.Body).Decode(&folders); err != nil {
		return "", err
	}

	var parentFolderID string
	for _, folder := range folders.Value {
		if folder.Name == parentFolderName {
			parentFolderID = folder.ID
			break
		}
	}

	if parentFolderID == "" {
		return "", fmt.Errorf("parent folder %s not found", parentFolderName)
	}

	// Step 2: Get the child folders of the parent folder (e.g., "Inbox")
	childFolderURL := fmt.Sprintf("%s/users/%s/mailFolders/%s/childFolders", GraphAPI, userID, parentFolderID)
	req, err = http.NewRequest("GET", childFolderURL, nil)
	if err != nil {
		return "", err
	}
	req.Header.Set("Authorization", "Bearer "+AccessToken)

	resp, err = http.DefaultClient.Do(req)
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		body, _ := ioutil.ReadAll(resp.Body)
		return "", fmt.Errorf("failed to get child folders: %s, %s", resp.Status, string(body))
	}

	var childFolders struct {
		Value []struct {
			ID   string `json:"id"`
			Name string `json:"displayName"`
		} `json:"value"`
	}
	if err := json.NewDecoder(resp.Body).Decode(&childFolders); err != nil {
		return "", err
	}

	for _, folder := range childFolders.Value {
		if folder.Name == targetFolderName {
			return folder.ID, nil
		}
	}

	return "", fmt.Errorf("folder %s not found in %s", targetFolderName, parentFolderName)
}
func flagEmail(userID, messageID string) error {
	url := fmt.Sprintf("%s/users/%s/messages/%s", GraphAPI, userID, messageID)

	// Request body to flag the email
	requestBody := map[string]interface{}{
		"flag": map[string]string{
			"flagStatus": "flagged",
		},
	}

	requestData, err := json.Marshal(requestBody)
	if err != nil {
		return err
	}

	req, err := http.NewRequest("PATCH", url, bytes.NewBuffer(requestData))
	if err != nil {
		return err
	}
	req.Header.Set("Authorization", "Bearer "+AccessToken)
	req.Header.Set("Content-Type", "application/json")

	resp, err := http.DefaultClient.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		body, _ := ioutil.ReadAll(resp.Body)
		return fmt.Errorf("failed to flag email: %s, %s", resp.Status, string(body))
	}

	log.Printf("Email with ID %s flagged successfully.\n", messageID)
	return nil
}

func monitorFolder(userID, folderID string) error {
	url := fmt.Sprintf("%s/users/%s/mailFolders/%s/messages?$filter=(flag/flagStatus eq 'notFlagged')", GraphAPI, userID, folderID)
	req, err := http.NewRequest("GET", url, nil)
	if err != nil {
		return err
	}
	req.Header.Set("Authorization", "Bearer "+AccessToken)

	resp, err := http.DefaultClient.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		body, _ := ioutil.ReadAll(resp.Body)
		return fmt.Errorf("failed to get messages: %s, %s", resp.Status, string(body))
	}
	// body, _ := ioutil.ReadAll(resp.Body)
	// fmt.Println(string(body))
	var messages struct {
		Value []Email `json:"value"`
	}
	if err := json.NewDecoder(resp.Body).Decode(&messages); err != nil {
		return err
	}

	for _, message := range messages.Value {
		// Compose and send a new email
		subject := "Re: " + message.Subject
		body := fmt.Sprintf("Thank you for your email. Here's a reply:\n\n%s", message.Body.Content)
		recipient := message.Sender.EmailAddress.Address

		if err := sendNewEmail(userID, recipient, subject, body); err != nil {
			log.Printf("Error sending email: %v\n", err)
			continue
		}
		log.Printf("Email sent successfully to %s\n", recipient)

		// Flag the email
		if err := flagEmail(userID, message.ID); err != nil {
			log.Printf("Error flagging email: %v\n", err)
		}
	}

	return nil
}

func sendNewEmail(userID, recipient, subject, body string) error {
	url := fmt.Sprintf("%s/users/%s/sendMail", GraphAPI, userID)

	requestBody := map[string]interface{}{
		"message": map[string]interface{}{
			"subject": subject,
			"body": map[string]string{
				"contentType": "Text",
				"content":     body,
			},
			"toRecipients": []map[string]interface{}{
				{
					"emailAddress": map[string]string{
						"address": recipient,
					},
				},
			},
		},
		"saveToSentItems": "true",
	}

	requestData, err := json.Marshal(requestBody)
	if err != nil {
		return err
	}

	req, err := http.NewRequest("POST", url, bytes.NewBuffer(requestData))
	if err != nil {
		return err
	}
	req.Header.Set("Authorization", "Bearer "+AccessToken)
	req.Header.Set("Content-Type", "application/json")

	resp, err := http.DefaultClient.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusAccepted {
		body, _ := ioutil.ReadAll(resp.Body)
		return fmt.Errorf("failed to send email: %s, %s", resp.Status, string(body))
	}

	return nil
}

func main() {
	if err := getAccessToken(); err != nil {
		log.Fatalf("Error obtaining access token: %v\n", err)
	}

	userID := "finance@remotesupportnederland.nl"
	parentFolderName := "Inbox"
	targetFolderName := "TestFolder"

	folderID, err := getFolderID(userID, parentFolderName, targetFolderName)
	if err != nil {
		log.Fatalf("Error getting folder ID: %v\n", err)
	}

	if err := monitorFolder(userID, folderID); err != nil {
		log.Fatalf("Error monitoring folder: %v\n", err)
	}

	log.Println("Emails processed successfully.")
}
