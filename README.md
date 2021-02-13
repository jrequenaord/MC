# MC
Description:
CM-API is a project which aims to automatise the creation of a report including manual data entries related with menstrual cycle and weight. On the one hand, some data such as initial date, basal temperature (Temperature (°C)), cervical mucus type (MC), quantity of menstrual blood (M-Quantity) are user inputs collected by an API hosted on a cloud instance. On the other hand, the weight measurements are picked up directly from Withings public API. This requires to have a Withings' scale. The objective of including weight on the report is to be able to compare it between the same phase of different cycles (i.e., a week after menstruation).
When the user inputs the data, an email which contains the report is sent to the configured email address. The whole process takes around ten seconds. 

In order to reduce the cost of hosting the cloud instance, it was decided to not use a static IP. For this reason, each morning the user is notified with the new IP where the API is available. This feature will be changed in future iterations.


Next steps:
- To grant a static IP to the instance hosting the API.
- To personalise the email's body.
- To improve the application's front end.
- To add log in and sign in capabilities.
- To add historic features.
- To be able to perform comparisons between cycles.
- To dockerise the application.
- To refactor the code.
- To modify MC values according to future knowledge acquired.
