# Service Board Imports

## Overview
This folder contains Excel exports from the F2 Service Board software, representing completed work from the service team.

## File Naming Convention
Files follow a consistent naming pattern:
- Format: `Service [date] at [time].xlsx`
- Example: `Service 2025-12-03 at 3.41.39 PM.xlsx`

## File Structure
The Excel file format is consistent across all imports. The structure should never change.

### Headers (Row 1)
1. **PrepDate** - Date when preparation was scheduled/completed
2. **ServicePriority_aet** - Priority level (e.g., ASAP, High)
3. **AssetBarcode** - Barcode identifier for the asset
4. **EquipmentName_lu** - Name of the equipment
5. **EquipmentCategory_lu** - Category classification
6. **OrderNumber_lu** - Order number reference
7. **JobName_lu** - Name of the job/project
8. **Puller_lu** - Person who pulled the equipment
9. **PrepTech_lu** - Prep technician assigned
10. **ServiceTech** - Service technician assigned
11. **EstimatedCompletionTime_t** - Estimated completion time
12. **z_log_CreateHost_ts_ae** - Timestamp when record was created
13. **TimestampStart_ts** - Start timestamp
14. **TimestampEnd_ts** - End timestamp
15. **TimestampDuration_cti** - Duration of service
16. **ServiceStatus_ct** - Current status of the service
17. **ServiceNotes** - Additional notes

## Workflow
1. F2 software exports file to user's Desktop
2. User manually copies/drags file into this folder
3. Apps Script processes the file (manual trigger or scheduled)
4. Data is imported, analyzed, and integrated into the system

## Notes
- File format should remain consistent across all exports
- Files are manually moved from Desktop to this folder
- Processing can be triggered manually or automatically via scheduled triggers

