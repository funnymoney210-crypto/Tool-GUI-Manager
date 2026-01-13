# Tool-GUI-Manager
Design and implement a new graphical user interface (GUI) to manage and control the tool's existing GUI.
Overview:  
The Tool-GUI-Manager is a desktop-based application designed to provide a simple, intuitive interface for managing and controlling tool-related operations. The application supports multiple languages (with a focus on Hebrew and English) and run independently on any PC within the network.

---

Implemented Features:

1. Multi-Language Support  
   - The interface and all input fields were implemented to support both Hebrew and English.  
   - Users can select their preferred language, which is saved and loaded automatically in future sessions.

2. User Interface  
   - A clean and user-friendly layout was designed, enabling quick navigation and clear data entry.  
   - The application is portable and can run from any PC on the network without installation.  
   - Input fields were clearly labeled, and tooltips were added to assist users.

3. Data Validation & Input Handling  
   - Real-time validation was implemented for all input fields to ensure data completeness and correct formatting.
   - Users receive instant feedback for missing or incorrect entries, such as invalid Work Order numbers or missing Cu thickness values.  
   - The system prevents form submission until all required inputs are validated.

4. Recipe Creation Logic  
   - Users are able to generate process recipes based on the number of wafers and Cu thickness.  
   - The system calculates optimal Dump & Refresh times using predefined formulas to ensure etch rate (ER) stability.

5. Etch Rate (ER) Monitoring and Control  
   - An SPC-based monitoring system was integrated to track ER in real-time.  
   - ER trends are displayed dynamically in chart format.  
   - If the ER exceeds control limits (OOC), the system blocks the user from proceeding and displays a clear warning message.

     
6. Cu Thickness to Etch Time Calculation  
   - A formula-based logic was embedded to derive the required etch time from the measured Cu thickness.  
   - A manual override option is available for authorized personnel.

---
