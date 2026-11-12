# Repository Guidance for Agents

## Downstream Office Wrappers

When changing COM wrapper APIs, behavior, or the application lifecycle in this repository, review the downstream Office wrappers that use them:

- [CompuMaster.MsAccess](https://github.com/CompuMasterGmbH/CompuMaster.MsAccess), including its AccessApplication wrapper.
- [CompuMaster.Excel](https://github.com/CompuMasterGmbH/CompuMaster.Excel), including both MsExcelComInterop and ExcelOps-MicrosoftExcel.

Update affected wrappers and their relevant tests where needed. If a downstream change is required for compatibility, correct behavior, or acceptance of the shared wrapper change, coordinate it as an integral part of the work, using linked pull requests in the affected repositories. For downstream work that can be handled separately, open linked issues in the affected repositories where appropriate. Describe the impact and the expected update in each issue.

If it is unclear whether downstream changes should be included in the current work or tracked separately, ask the developer whether to include them before expanding the scope.
