{
  "input_csv": "input_file.csv",
  "output_excel": "output_step5.xlsx",

  "columns_to_remove": [
    "DummyColumn1", "DummyColumn2", "DummyColumn3",
    "DummyColumn4", "DummyColumn5", "DummyColumn6",
    "DummyColumn7", "DummyColumn8", "DummyColumn9"
  ],

  "columns_to_add": [
    "Description",
    "Remediation Steps",
    "Environment",
    "Primary Contact",
    "Manager / Sr Manager / Director / Sr Director",
    "Sr Director / VP",
    "VP / SVP / CVP",
    "BU"
  ],

  "parse_account_column": true,
  "account_column_name": "Account",
  "resource_column_name": "Resource ID",

  "mappings": [
    {
      "name": "Remediation Mapping",
      "file": "Report_Anex.xlsx",
      "sheet": "Anex1_Remediation_Sheet",
      "join_column": "Policy ID",
      "source_to_target": {
        "Policy Statement": "Description",
        "Policy Remediation": "Remediation Steps"
      }
    },
    {
      "name": "Subscription Mapping",
      "file": "Report_Anex.xlsx",
      "sheet": "Anex2_Sub_Sheet",
      "join_column": "Subscription ID",
      "source_to_target": {
        "Environment": "Environment",
        "Primary Contact": "Primary Contact"
      }
    }
  ]
}
