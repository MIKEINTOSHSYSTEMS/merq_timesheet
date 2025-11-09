MERQ TIMESHEET APPLICATION
==========================

OVERVIEW
--------
The MERQ Timesheet Application is a comprehensive work hour tracking system 
specifically designed for MERQ Consultancy PLC employees and consultants. 
This application facilitates accurate time tracking using the Ethiopian 
calendar system with professional reporting capabilities.

Objective 
---------
The main objective of this application is to generate a formatted excel documentable Timesheet for Monthly submission!


Application Version
-------------------

Version: 1.0
Developer: Information Systems & Digital Health Unit (ISDHU)
Company: MERQ Consultancy PLC
Website: https://merqconsultancy.org/

KEY FEATURES
------------

1. ETHIOPIAN CALENDAR INTEGRATION
   • Accurate Ethiopian date conversion algorithms
   • Real-time Ethiopian calendar display
   • Support for all 13 Ethiopian months including Pagume
   • Bilingual interface (Amharic and English)

2. PROJECT TIME TRACKING
   • Multiple project management
   • Daily hour allocation per project
   • Project-wise hour allocation limits
   • Automatic total hour calculations
   • Project summary and analytics

3. LEAVE AND ABSENCE MANAGEMENT
   • Vacation leave tracking
   • Sick leave recording
   • Holiday and special leave
   • Personal leave management
   • Bereavement leave
   • Other leave categories

4. EXCEL EXPORT FUNCTIONALITY
   • Professional Excel template integration
   • Automated timesheet generation
   • Formatted reports with MERQ branding
   • Signature sections for approval
   • Watermark and security features

5. USER-FRIENDLY INTERFACE
   • Modern, professional design
   • Bilingual support (Amharic/English)
   • Real-time validation and error checking
   • Scrollable calendar interface
   • Responsive layout design

6. DATA VALIDATION AND SECURITY
   • Input validation for hour entries
   • 24-hour daily limit enforcement
   • Future date prevention
   • Data integrity checks
   • Confidentiality safeguards

SYSTEM REQUIREMENTS
-------------------

MINIMUM SYSTEM REQUIREMENTS:
• Operating System: Windows 7 or later
• Processor: 1 GHz or faster
• Memory: 2 GB RAM
• Storage: 100 MB free disk space
• Display: 1024x768 resolution

RECOMMENDED SYSTEM REQUIREMENTS:
• Operating System: Windows 10 or later
• Processor: 2 GHz or faster
• Memory: 4 GB RAM or more
• Storage: 200 MB free disk space
• Display: 1366x768 resolution or higher

SOFTWARE DEPENDENCIES:
• Python 3.8 or higher (included in installer)
• Required Python packages (automatically installed):
  - pandas
  - openpyxl
  - Pillow (PIL)
  - requests

INSTALLATION INSTRUCTIONS
-------------------------

AUTOMATIC INSTALLATION:
1. Run "MERQ_Timesheet_Setup.exe"
2. Follow the installation wizard prompts
3. Choose installation directory (default recommended)
4. Select desired options:
   - Create desktop shortcut (recommended)
   - Pin to taskbar (Windows 10/11)
5. Complete installation
6. Launch the application from Start Menu or desktop

MANUAL INSTALLATION (Development):
1. Ensure Python 3.8+ is installed
2. Install required packages:
   pip install pandas openpyxl pillow requests
3. Run the application:
   python MERQ_Timesheet.py

USAGE GUIDE
-----------

GETTING STARTED:
1. Launch the MERQ Timesheet Application
2. Read and accept the disclaimer
3. Enter your full name in the designated field
4. Select the Ethiopian year and month

ADDING PROJECTS:
1. Click "Add Project" button
2. Enter project name
3. Set allocated hours (if known)
4. Repeat for additional projects

ENTERING WORK HOURS:
1. Navigate to the timesheet grid
2. For each project, enter hours worked per day
3. Use "Prefill Default Hours" for standard work patterns
4. System automatically validates and calculates totals

MANAGING LEAVE:
1. Scroll to the leave section
2. Select appropriate leave type
3. Enter leave hours for relevant days
4. System tracks leave balances

EXPORTING TIMESHEET:
1. Click "Preview Timesheet Summary" to review
2. Verify all entries are correct
3. Click "Download Timesheet Excel" to export
4. Save the file with your name and month
5. Submit to supervisor as required

VALIDATION RULES
----------------

• Maximum 24 hours per day across all projects and leave
• Only numeric values accepted for hour entries
• Future dates cannot be selected
• Employee name must be entered before data entry
• All fields are validated in real-time

TROUBLESHOOTING
---------------

COMMON ISSUES AND SOLUTIONS:

1. APPLICATION WON'T START:
   • Ensure Windows is updated
   • Run as Administrator
   • Check antivirus software isn't blocking the application

2. EXCEL EXPORT FAILS:
   • Ensure template file is present
   • Check write permissions in save location
   • Verify Microsoft Excel is not running in background

3. CALENDAR DISPLAY ISSUES:
   • Verify system date and time settings
   • Check internet connection for date API fallback

4. DATA VALIDATION ERRORS:
   • Ensure total hours don't exceed 24 per day
   • Check for non-numeric characters in hour fields
   • Verify all required fields are completed

SUPPORT INFORMATION
-------------------

TECHNICAL SUPPORT:
Email: support@merqconsultancy.org
Phone: +251913391985
Hours: 8:30 AM - 5:30 PM EAT, Monday to Friday

REPORTING ISSUES:
When contacting support, please provide:
• Application version number
• Windows version
• Detailed description of the issue
• Screenshots if possible
• Steps to reproduce the problem

DATA BACKUP AND RETENTION
-------------------------

• Timesheet data is saved in Excel format
• Users should maintain their own backups
• Submitted timesheets become official records
• Data retention follows MERQ Consultancy policies

SECURITY AND PRIVACY
--------------------

• Application is for internal MERQ use only
• All data remains confidential to MERQ
• No external data transmission except date API
• Local installation ensures data privacy

UPDATES AND MAINTENANCE
-----------------------

• Regular updates will be provided as needed
• Users will be notified of new versions
• Update instructions will be provided with releases
• Back up data before updating

LEGAL AND COMPLIANCE
--------------------

COPYRIGHT:
© 2025 MERQ Consultancy PLC. All Rights Reserved.

LICENSE:
Proprietary software for internal MERQ use only.

DISCLAIMER:
This software is provided "AS IS" without warranties.
Users are responsible for accurate data entry.

DEVELOPER INFORMATION
---------------------

Development Team: Information Systems & Digital Health Unit (ISDHU)
Lead Developer: ISDHU Team
Quality Assurance: MERQ Consultancy PLC
Testing: Internal MERQ testing team

SPECIAL NOTES
-------------

• This application uses Ethiopian calendar calculations
• All times are based on Ethiopian working hours
• Reports are formatted for MERQ administrative requirements
• Application is optimized for MERQ business processes

VERSION HISTORY
---------------

Version 1.0 (November 2025)
• Initial release
• Ethiopian calendar integration
• Excel template export
• Bilingual interface

FREQUENTLY ASKED QUESTIONS
--------------------------

Q: Can I use this application for multiple months?
A: Yes, simply change the month and year selection.

Q: What if I make a mistake in my entries?
A: You can clear individual entries or use "Clear All" to start over.

Q: Is my data saved automatically?
A: Data is saved when you export to Excel. Keep backups.

Q: Can I use this on multiple computers?
A: Installation is per computer. You can export and import data via Excel.

Q: What if the Ethiopian date appears incorrect?
A: The application uses verified algorithms. Contact support if issues persist.

CONTACT FOR FEEDBACK
--------------------

We welcome feedback to improve this application. Please send suggestions to:
feedback@merqconsultancy.org

--------------------------------------------------------------------
THANK YOU FOR USING MERQ TIMESHEET APPLICATION!
--------------------------------------------------------------------

This application was developed to streamline time tracking and improve
efficiency for all MERQ Consultancy PLC team members.

Remember to submit your completed timesheets promptly at the end of each month.

For urgent matters, contact your supervisor or the ISDHU team directly.

Last Updated: November 2025