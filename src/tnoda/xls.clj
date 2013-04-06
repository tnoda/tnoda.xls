(ns tnoda.xls
  (:import (org.apache.poi.ss.usermodel Cell Row)
           (org.apache.poi.poifs.filesystem POIFSFileSystem)
           (org.apache.poi.hssf.usermodel HSSFWorkbook HSSFSheet)
           (java.io FileInputStream)))

(defn- cell-type
  [^Cell cell]
  (if-let [t (and cell (.getCellType cell))]
    (if (= t Cell/CELL_TYPE_FORMULA)
        (.getCachedFormulaResultType cell)
        t)))

(defn- unsupported-cell-type
  [^Cell cell]
  (throw (IllegalArgumentException. (str "Unsupported cell type at ("
                                        (.getRowIndex cell)
                                        ","
                                        (.getColumnIndex cell)
                                        "): "
                                        (cell-type cell)))))

(defn- cell->value
  [^Cell cell]
  (if-let [t (cell-type cell)]
    (cond (= t Cell/CELL_TYPE_BOOLEAN) (.getBooleanCellValue cell)
          (= t Cell/CELL_TYPE_NUMERIC) (.getNumericCellValue cell)
          (= t Cell/CELL_TYPE_STRING)  (.getStringCellValue  cell)
          :default (unsupported-cell-type cell))))

(defn- row->values
  [^Row row]
  (->> (.getLastCellNum row)
       range
       (map (comp cell->value
                  #(.getCell row % Row/RETURN_BLANK_AS_NULL)))))

(defn- sheet ^HSSFSheet
  [^HSSFWorkbook wb idx]
  (cond (number? idx) (.getSheetAt wb (int idx))
        (string? idx) (.getSheet wb (str idx))
        :default (throw (IllegalArgumentException. (str idx " is neither number nor string.")))))

(defn read-sheet
  "Reads a sheet at the given idx, and returns a lazy seq of rows.
Each row is represented as a lazy seq of cells. If a cell is either
empty or blank, its value will be nil. idx should be either the
sheet number of the sheet name. If idx is omitted, it is set to be
zero."
  ([^FileInputStream fis]
     (read-sheet fis 0))
  ([^FileInputStream is idx]
     (let [sheet (-> is POIFSFileSystem. HSSFWorkbook. (sheet idx))]
       (->> sheet
            .rowIterator
            iterator-seq
            (map row->values)))))
