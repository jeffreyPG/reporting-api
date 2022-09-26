# reporting-api

Information about this API can be found here: https://github.com/simuwatt/node-api/wiki/Reporting#c-api



The solution holds 5 projects at the moment:

  1. ChartsAPI is still living in this solution and can be deleted once https://github.com/simuwatt/chart-api is deployed and tested thoroughly.
  2. HtmlToOpenXml is a class library project that we use to convert HTML to OpenXML readable by Word documents. It was initially built from https://github.com/onizet/html2openxml
  3. reports is the current reports-api that is deployed in both QA and beta environments.
  4. reports.tests is the test project that we can use on automated deployments as a test step to make sure the methods still works.
  5. ReportsAPISite is the WebApp version of reports-api. It is not being used yet.
