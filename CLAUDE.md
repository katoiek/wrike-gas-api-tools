# CLAUDE.md

必ず日本語で回答してください。
This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Google Apps Script (GAS) project that provides comprehensive tools for interacting with the Wrike API. The main functionalities include importing data from Backlog into Wrike, user management, and various administrative operations.

## Architecture

### Core Components

- **main.gs**: Central authentication system with OAuth2 flow, script properties management, and UI menu system
- **ImportBacklogToWrike.gs**: Main import functionality for converting Backlog data to Wrike tasks
- **BulkInviteUsers.gs**: Bulk user invitation system with role management and spreadsheet integration
- **UpdateUsertoCollaborator.gs**: Batch user role conversion and group management
- **GetUserIdv4.gs**: API v2 to v4 user ID conversion utility
- **GetSpaceList.gs**: Retrieves Wrike spaces/folders data
- **GetAllContacts.gs**: Retrieves contact information from Wrike
- **GetFolderProjectList.gs**: Gets folder/project listings
- **GetAllCustomFields.gs**: Retrieves custom field definitions
- **Get Information about me.gs**: Retrieves current user information from Wrike API

### Common Patterns

1. **OAuth2 Authentication**: Uses OAuth2.createService() for token management with automatic refresh
2. **Script Properties Management**: Configuration stored in `parameters` sheet (B1-B11) with `registKeys()` function
3. **API Rate Limiting**: All API calls include `Utilities.sleep(100-500)` for rate limit compliance
4. **Batch Processing**: Functions process data in batches (typically 10 items) to handle large datasets
5. **Error Handling**: Comprehensive try-catch blocks with user-friendly error messages via `SpreadsheetApp.getUi().alert()`
6. **Sheet Management**: Functions follow a pattern of clearing existing data, setting headers, and populating results in named sheets

### Authentication Flow

1. Parameters are stored in `parameters` sheet and loaded via `registKeys()`
2. OAuth2 service is configured with client credentials and callback URL
3. `onOpen()` automatically checks token status and prompts for re-authentication
4. Token is stored in script properties after successful authentication

### Data Flow

1. **Backlog Import**: Data stored in `BacklogData` sheet → B1 cell contains Wrike folder permalink → Processing extracts folder ID → Comments (columns 1-20) are concatenated → Creates Wrike tasks
2. **User Management**: User data in `Userlist` sheet → ID conversion from v2 to v4 → Role updates → Group assignments
3. **Bulk Invitations**: User data in `BulkUserInvite` sheet → Role validation → Batch invitation processing → Results written back to sheet

## Configuration Requirements

### Script Properties (via parameters sheet)
- B1: `api_url` - Wrike API base URL
- B2: `client_id` - OAuth2 client ID
- B3: `client_secret` - OAuth2 client secret
- B4: `token_url` - OAuth2 token endpoint
- B5: `auth_url` - OAuth2 authorization endpoint
- B6: `scope` - OAuth2 scope
- B7: `account_Id` - Wrike account ID
- B11: Callback URL (auto-generated)

### Spreadsheet Structure
- **parameters**: Configuration data (B1-B11)
- **BacklogData**: Import data with headers in row 2, data starting from row 3
- **BulkUserInvite**: User invitation data with role dropdowns
- **Userlist**: User ID management (A column: v4 IDs, B column: v2 IDs)
- Various output sheets for API results

## Development Commands

### Testing Functions
```javascript
// Test single task creation
testCreateSingleTask()

// Validate configuration
checkConfiguration()

// Check token status
checkTokenStatus()
```

### Main Operations
```javascript
// Initialize bulk invite sheet
initBulkUserInviteSheet()

// Execute bulk user invitation
BulkInviteUsers()

// Update users to collaborator role
UpdateUsertoCollaborator()

// Convert user IDs from v2 to v4
GetUserIdv4()
```

## Key Functions

### Authentication & Setup
- `registKeys()`: Loads configuration from parameters sheet
- `showAuth()`: Initiates OAuth2 authentication
- `checkTokenStatus()`: Validates current token
- `onOpen()`: Auto-authentication and menu setup

### Import Functions
- `ImportBacklogToWrike()`: Main import function
- `extractFolderIdFromPermalink()`: Converts permalink to folder ID
- `createWrikeTask()`: Creates individual Wrike tasks

### User Management
- `BulkInviteUsers()`: Bulk user invitation with role assignment
- `UpdateUsertoCollaborator()`: Batch role conversion
- `GetUserIdv4()`: ID conversion utility
- `initBulkUserInviteSheet()`: Sets up invitation spreadsheet

### API Data Retrieval
- `getSpaceList()`: Gets Wrike spaces
- `GetAllContacts()`: Gets all contacts
- `getInformationAboutMe()`: Gets current user info

## Development Notes

- All code includes comprehensive Japanese comments and error messages
- API calls implement proper rate limiting (100-500ms delays)
- Batch processing handles large datasets efficiently
- OAuth2 flow provides automatic token refresh
- UI includes custom menu system with numbered workflow items
- Functions include proper error handling and user feedback through modeless dialogs
- Some files contain commented-out alternative implementations for different API approaches
