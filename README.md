# oxml-doc
Library for streaming writing/reading xlsx files;

---
# Build status:
| Branch | Status |
|--------|--------|
| master |[![Build Status](https://travis-ci.org/protei-rnd/oxml-doc.svg?branch=master)](https://travis-ci.org/protei-rnd/oxml-doc)|

---

# How to use:
```java
WorkSheet sheet = writer.createNewSheet("sheet");
Row row = new Row();
for (int i = 0; i < 100000; i++) {
    for (int j = 0; j < 25; j++)
        row.append("aaaaa");
    sheet = sheet.appendRow(row);
    row.clear();
}
final File bigDoc = new File("large_doc.xlsx");
writer.writeTo(bigDoc);
```
