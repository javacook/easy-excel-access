# easy-excel-access
An API based on POI to access Excel content in an easy "functional" way.

## Installation
### Maven
```
<dependency>
  <groupId>de.javacook</groupId>
  <artifactId>easy-excel-access</artifactId>
  <version>{version}</version>
</dependency>
```
### Gradle
```
compile group: 'de.javacook', name: 'easy-excel-access', version: '{version}'
```

### Hint
Version 0.8 is based on Apache POI 3.14, version 0.9.* on Apache POI 4.1.1. 
These two POI versions are not API compatible! 

## Code snippets
```
new ExcelCoordinateSequencer()
        .fromCol(3).fromRow(5).toCol(5).toRow(6)
        .sequence()
        .forEach(System.out::println);
```
```
new ExcelCoordinateSequencer()
        .fromCol(3).fromRow(5).toCol(5).toRow(6)
        .forEach((coord, i) -> System.out.println(coord.col() + " - " + i));
```
```
new ExcelCoordinateSequencer()
        .from(2, 3).to(7, 9)
        .stopWhen(coord -> coord.col() + coord.row() > 10)
        .forEach(System.out::println);
```
```
new ExcelCoordinateSequencer()
        .from(2, 3).to(7, 9)
        .stopWhen(coord -> coord.hasSameCol(new ExcelCoordinate(4,8)))
        .forEach(System.out::println);
```
```
new ExcelCoordinateSequencer()
        .from(2, 3).to(7, 9)
        .stopWhen(coord -> coord.hasSameRow(new ExcelCoordinate(4,8)))
        .forEach(System.out::println);
```
```
new ExcelCoordinateSequencer()
        .from(excel.find("Abfrageperiode").addRow(2).setCol("D"))
        .height(3).width(1)
        .sequence()
        .forEach(System.out::println);
```