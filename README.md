# gen_interconnection_queues

Author: Selorm Kwami Dzakpasu

Date: 1/13/2025

Topic: Compilation of Generator Interconnection Queues from the Seven ISOs (CAISO, ERCOT, MISO, PJM, SPP, NYISO & NEISO)

Sources(ISOs):

*CASIO - https://www.caiso.com/library/recent-documents

*ERCOT - https://www.ercot.com/mp/data-products/data-product-details?id=PG7-200-ER

*MISO - https://www.misoenergy.org/planning/resource-utilization/GI_Queue/gi-interactive-queue/

*PJM - https://www.pjm.com/planning/service-requests/serial-service-request-status

*SPP - https://opsportal.spp.org/Studies/GIActive

*NYISO - https://www.nyiso.com/documents/20142/1407078/NYISO-Interconnection-Queue.xlsx

*NEISO - https://irtt.iso-ne.com/reports/external


Things to note:

* Included in this folder are Python scripts to process the queues from the above listed ISOs and an additional script to compile the processed ISO queues into one large Excel file/database.

* Some functions within the Python scripts are not applicable to .csv files. This is indicated in the comments, and unsupported file types will result in an error message. 

* If you encounter an error while running any of the scripts to process a queue in .csv format, first convert the file to .xlsx, and you should be good to go!

* To use this resource:
1. First, download the latest interconnection queues from the sources listed above and store them in the "Queues" folder.
2. Change the file paths within the individual Python scripts created for each ISO to reflect the file names of the newly downloaded queues. NB: There is no need to update "Compiler.py"
3. Run only the "main.py" Python script.

Enjoy!
