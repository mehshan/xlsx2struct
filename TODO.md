# TODO

- [ ] Test freely available xlsx sheets
- [ ] Test xlsx files created with different versions Excel (e.g. office 365)
- [ ] Test write xlsx and read it back
- [ ] Add docs in readme
- [ ] New xlsx2struct repo, GitHub CI/CD and releases
- [ ] Complete code documentation
- [ ] Document and test unsupported features (e.g. nested structures)

### In Progress

- [ ] Complete all test cases and code coverage

### Done ✓

- [x] Setup repo structure
- [x] Re-factor field.go tests
- [x] Support for system types and other commonly used types (e.g. time)
- [x] Field options in tags `column:"heading=Order Date,trim,time=2006-01-01"`
- [x] Handle empty rows in unmarshalFields
- [x] Revise reflection code in newStruct
