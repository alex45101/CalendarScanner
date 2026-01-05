# CalendarScanner

CalendarScanner is a C# .NET 8 application that automates the process of extracting work schedules from email image attachments using Azure OCR, and creates corresponding events in a Google Calendar. It also sends a confirmation email upon completion.

## Features

- Connects to a Gmail inbox via IMAP to find unread "Work schedule" emails.
- Downloads the first image attachment from the relevant email.
- Uses Azure Computer Vision OCR to extract text from the image.
- Parses the recognized text into a structured work schedule.
- Creates Google Calendar events for each workday.
- Sends a confirmation email after processing.

## Prerequisites

- .NET 8 SDK
- Visual Studio 2022 or later
- Azure Computer Vision API key and endpoint
- Google Calendar API credentials
- Gmail account with IMAP enabled and an app password for authentication

## Configuration

The application requires a configuration file or setup method to provide the following information (see `SetupInfo`):

- AzureEndpoint: Azure Computer Vision endpoint URL
- AzureKey: Azure Computer Vision API key
- ScannerEmail: Gmail address to scan
- ScannerEmailAPIPassword: App password for Gmail
- CalendarKey: Google Calendar API key
- ApplicationName: Name for Google Calendar API
- SendConfirmation: Whether to send a confirmation email
- ConfirmationName: Name for confirmation email recipient
- ConfirmationEmail: Email address for confirmation
- ConfirmationSubject: Subject for confirmation email
- ConfirmationBody: Body for confirmation email
- EventName: Name for created calendar events
- EventLocation: Location for events
- EventTimeZone: Time zone for events
- ScheduleFormat: Path to JSON file describing the schedule format
- LogFilePath: Path to log file

## Usage

1. **Configure** your `SetupInfo` (see above).
2. **Run the application**:
