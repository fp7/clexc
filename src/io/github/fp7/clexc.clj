(ns io.github.fp7.clexc
  (:require [clojure.java.io :as io])
  (:import (org.apache.poi.ss.usermodel BuiltinFormats)
           (org.apache.poi.ss.usermodel Cell)
           (org.apache.poi.ss.usermodel CellType)
           (org.apache.poi.ss.usermodel DateUtil)
           (org.apache.poi.ss.usermodel Row)
           (org.apache.poi.ss.usermodel Sheet)
           (org.apache.poi.ss.usermodel Workbook)
           (org.apache.poi.xssf.usermodel XSSFWorkbook)))

(set! *warn-on-reflection* true)

(defn ^:private set-value
  [^Cell cell cell-value]
  (cond
    (string? cell-value) (.setCellValue cell ^String cell-value)
    (number? cell-value) (.setCellValue cell (double cell-value))
    (boolean? cell-value) (.setCellValue cell (boolean cell-value))
    (nil? cell-value) (.setBlank cell)
    (inst? cell-value) (let [wb (.. cell (getSheet) (getWorkbook))
                             cs (.createCellStyle wb)]
                         (doto cell
                           (.setCellValue (java.util.Date. (long (inst-ms cell-value))))
                           (.setCellStyle (doto cs
                                            (.setDataFormat
                                             (BuiltinFormats/getBuiltinFormat "m/d/yy h:mm"))))))
    :else (throw (ex-info "Value can not be set in cell" {:type (type cell-value)}))))

(defn ^:private add-row
  [^Sheet sheet ^long cnt row-data]
  (let [^Row row (.createRow sheet cnt)]
    (doseq [[cell-cnt  cell-value] (partition 2 (interleave (range) row-data))]
      (let [^Cell cell (.createCell row cell-cnt)]
        (set-value cell cell-value)))))

(defn ^:private add-sheet
  [^Workbook ws ^String sheet-name data]
  (let [sheet (.createSheet ws sheet-name)]
    (doseq [[cnt row] (partition 2 (interleave (range) data))]
      (add-row sheet cnt row ))))

(defn write-xlsx
  [p xlsx]
  (let [ws (XSSFWorkbook.)]
    (doseq [[sheet-name data] xlsx]
      (add-sheet ws sheet-name data))
    (with-open [c (io/output-stream p)]
      (.write ws c))))


(defn ^:private read-cell
  [^Cell cell]
  (cond (DateUtil/isCellDateFormatted cell) (.getDateCellValue cell)
        :else
        (condp = (.getCellType cell)
          CellType/STRING (.getStringCellValue cell)
          CellType/NUMERIC (.getNumericCellValue cell)
          CellType/BLANK nil
          CellType/BOOLEAN (.getBooleanCellValue cell))))

(defn ^:private read-row
  [^Row row]
  (into []
        (map read-cell)
        (seq row)))

(defn ^:private read-sheet
  [^Sheet sheet]
  [(.getSheetName sheet)
   (into []
         (map read-row)
         (seq sheet))])

(defn read-xlsx
  [p]
  (with-open [c (io/input-stream p)]
    (let [ws (XSSFWorkbook. c)]
      (into {}
            (map read-sheet)
            (seq ws)))))


(comment
  (read-xlsx "foo.xlsx")
  )

(comment
  (write-xlsx "foo.xlsx" {"My sheet" [["hello world" "moin moin"] ["foo"]]
                          "foobar" [[]]})
  )
