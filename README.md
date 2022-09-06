Apache POI를 통한 Excel 파일 Write/Read 기능 구현
-

* https://www.baeldung.com/java-microsoft-excel 문서를 참고했습니다.

Dependency 추가
- poi -> 97-2003 legacy excel file (binary file format) 대응
- poi-ooxml -> 2007+ excel file (xml-based file format) 대응
- 참고 문서
  - https://stackoverflow.com/questions/60217698/what-is-the-difference-between-poi-and-poi-ooxml

Apache POI 기본설명
- Excel 파일을 모델링하기 위한 Workbook 인터페이스 제공
- Excel 파일의 요소를 모델링하는 Sheet, Row 및 Cell 인터페이스 제공
- .xls, .xlsx 파일 형식에 대한 각 인터페이스의 구현 제공
- 최신의 .xlsx 파일 형식으로 작업할 때 XSSFWorkbook, XSSFSheet, XSSFRow 및 XSSFCell 클래스 사용
- 이전 .xls 파일 형식으로 작업할 때 HSSFWorkbook, HSSFSheet, HSSFRow 및 HSSFCell 클래스 사용