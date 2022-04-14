# AzureAD-Guest-Account-Invite-Creator
 Sends out invitations for guest users in AzureAD

## HowTo ##
1) Open the script by exacuting it with the following 3 parameters:

| Script File | Absolute Path to Spreadsheet | Sheet you want to use |
|----------|----------|----------|
| .\GuestAccountInviteCreator.ps1 | "C:\Users\StefanKubisa\Documents\Scripts\GuestInvitationList.xlsx" | Sheet1 | 

Like so: 

.\GuestAccountInviteCreator.ps1 "C:\Users\StefanKubisa\Documents\Scripts\GuestInvitationList.xlsx" Sheet1

2) Make sure the static values match your worksheet's colum

| Status | Guest Name | Guest Email | 
|----------|----------|----------|
|  | Guest 1 | `Guest.1@domain.com` | 
|  | Guest 2 | `Guest.2@domain.com` | 
|  | Guest 3 | `Guest.3@domain.com` | 