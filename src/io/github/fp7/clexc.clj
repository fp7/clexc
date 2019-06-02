(ns io.github.fp7.clexc
  (:require [clojure.java.io :as io])
  (:import (org.apache.poi.ss.usermodel Cell)
           (org.apache.poi.ss.usermodel CellType)
           (org.apache.poi.ss.usermodel Row)
           (org.apache.poi.xssf.usermodel XSSFWorkbook)
           (org.apache.poi.xssf.usermodel XSSFSheet)))

(set! *warn-on-reflection* true)

(defn ^:private add-row
  [^XSSFSheet sheet ^long cnt row-data]
  (let [^Row row (.createRow sheet cnt)]
    (doseq [[cell-cnt  cell-value] (partition 2 (interleave (range) row-data))]
      (let [^Cell cell (.createCell row cell-cnt)]
        (.setCellValue cell ^String cell-value)))))

(defn ^:private add-sheet
  [^XSSFWorkbook ws ^String sheet-name data]
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
  (condp = (.getCellType cell)
    CellType/STRING (.getStringCellValue cell)))

(defn ^:private read-row
  [^Row row]
  (into []
        (map read-cell)
        (seq row)))

(defn ^:private read-sheet
  [^XSSFSheet sheet]
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
