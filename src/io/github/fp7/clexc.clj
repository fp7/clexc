(ns io.github.fp7.clexc
  (:require [clojure.java.io :as io])
  (:import (org.apache.poi.ss.usermodel Cell)
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

(comment
  (write-xlsx "foo.xlsx" {"My sheet" [["hello world" "moin moin"] ["foo"]]})
  )
