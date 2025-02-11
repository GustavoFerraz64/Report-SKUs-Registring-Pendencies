# Report PCI

This script aims to identify all PCIs (presumably, Purchase Change Initiatives or similar) that are currently in the registration workflow, determine the projects to which these PCIs are linked, and send emails to each department that has a pending PCI.

## Functionality

- Read the material codes from the "Código e campos.xlsx" spreadsheet.

- Enter the codes and fields in the ZM277 transaction to view the pending departments.

- Through the ZM277 transaction, access the SLA of these materials to obtain the entry date of this material in each department.

- Enter the materials in ZP059 to see which projects the materials are linked to.

- Assemble a consolidated table from the information of the two transactions.

- Send email to the responsible departments.

- Calculate how long the material has been in each pending department, using the information obtained in the SLA.

- Send email to DPCP management informing the times of each department.

@author: Gustavo Nunes Ferraz

@department: DPCP

## Modification History

- 2024-10-29: Script finalized.

- 2025-01-13: Access to the SLA of the PCIs + Sending email to the DPCP supervisor informing the times per department + Addition of the Readme and requirements.