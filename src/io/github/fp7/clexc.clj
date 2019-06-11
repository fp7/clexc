; Copyright (c) 2019 Finn Petersen
;
; This program and the accompanying materials are made
; available under the terms of the Eclipse Public License 2.0
; which is available at https://www.eclipse.org/legal/epl-2.0/
;
; SPDX-License-Identifier: EPL-2.0

(ns io.github.fp7.clexc
  (:require [clojure.java.io :as io])
  (:import (org.apache.poi.ss.usermodel Cell)
           (org.apache.poi.ss.usermodel CellType)
           (org.apache.poi.ss.usermodel DateUtil)
           (org.apache.poi.ss.usermodel Row)
           (org.apache.poi.ss.usermodel Sheet)
           (org.apache.poi.ss.usermodel Workbook)
           (org.apache.poi.xssf.usermodel XSSFWorkbook)))

(set! *warn-on-reflection* true)

(def ^:private DEFAULT_DATE_FORMAT "m/d/yy h:mm")

(defn ^:private set-value
  [^Cell cell cell-value]
  (let [cv (cond
             (map? cell-value) (:value cell-value)
             :else cell-value)
        wb (.. cell (getSheet) (getWorkbook))]
    (when-let [^String cf (or (:cell-format (meta cell-value))
                              (and (inst? cv)
                                   DEFAULT_DATE_FORMAT))]
      (.setCellStyle cell (doto (.createCellStyle wb)
                            (.setDataFormat (.. wb (createDataFormat) (getFormat cf))))))
    (cond
      (string? cv) (.setCellValue cell ^String cv)
      (number? cv) (.setCellValue cell (double cv))
      (boolean? cv) (.setCellValue cell (boolean cv))
      (nil? cv) (.setBlank cell)
      (inst? cv) (.setCellValue cell (java.util.Date. (long (inst-ms cv))))
      :else (throw (ex-info "Value can not be set in cell" {:type (type cv)})))))

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
  (let [cs (.. cell (getCellStyle) (getDataFormatString))
        m (when (not (#{DEFAULT_DATE_FORMAT "General"} cs))
            {:cell-format cs})
        v (cond (= CellType/STRING (.getCellType cell)) (.getStringCellValue cell)
                (= CellType/BLANK (.getCellType cell)) nil
                (= CellType/BOOLEAN (.getCellType cell)) (.getBooleanCellValue cell)

                (#{CellType/FORMULA CellType/_NONE CellType/ERROR} (.getCellType cell))
                (throw (ex-info "Don't know how to handle cell-type" {:cell-type (.getCellType cell)}))

                (DateUtil/isCellDateFormatted cell) (.getDateCellValue cell)
                (= CellType/NUMERIC (.getCellType cell)) (.getNumericCellValue cell))]
    (if (empty? m)
      v
      (with-meta {:value v} m))))

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
  (write-xlsx "foo.xlsx"  {"sheet 1" [[(with-meta
                                         {:value (java.util.Date.)}
                                         {:cell-format "h"})]]})
  )
