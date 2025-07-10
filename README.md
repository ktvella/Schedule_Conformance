# Schedule_Conformance
Tracking manufacturing floor conformance to the planned schedule for the week

One KPI of a manufacturing floor is how well it follows the planned schedule for the week. This script works with automatic exports from the ERP system at Hutchinson to track how well we are conforming to the schedule. When a department is not conforming to schedule, this report can be used to identify the reasons and drive corrective actions. 


Inputs: 
  - The automated exports for each weekday (e.g. "Monday Sched Conform WK23.csv")

Outputs: 
- Scheduled Manufacturing Orders for the week for each department (e.g. "DeptB Monday Scheduled MOs WK23")
  - Generated on Mondays. For visibility into what is on the schedule for the week 
- Plots of progress throughout the week of both % Hours complete and % Manufacturing Orders complete ("Status Week 23.png")
  - Generated every day. To keep track of progress  
- Excel file of progress ("Sch Conf Status WK23.xlsx")
  - same information as plots, but in table form 
- Reasons Sheet for each department of non-completed Manufacturing Orders ("DeptB Sch Conf Reasons WK23.xlsx")
  - Generated at the end of the week. For managers to fill out reasons for why Manufacturing Orders were not able to be completed
