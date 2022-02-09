package com.finskayaylochka.pdfreader;

import lombok.AccessLevel;
import lombok.Data;
import lombok.experimental.FieldDefaults;

/**
 * @author Alexandr Stegnin
 */
@Data
@FieldDefaults(level = AccessLevel.PRIVATE)
public class ExcelDocument {

  String address;
  String area;
  String sum;
  String numbers;
  String number;
  String owner;
  String info;

}
