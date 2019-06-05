(ns io.github.fp7.clexc-test
  (:require [clojure.test :as t]
            [io.github.fp7.clexc :as clexc])
  (:import (org.apache.poi.ss.usermodel CellType)))


(defn ^:private write-and-reread
  [example]
  (let [out-stream (java.io.ByteArrayOutputStream.)
        out (clexc/write-xlsx out-stream example)
        read (clexc/read-xlsx (java.io.ByteArrayInputStream. (.toByteArray out-stream)))]
    read))

(t/deftest cell-type-check
  (t/testing "Canary test for checking if new cell types are added. You have to revisit the read-cell method and adopt any changes so that numeric and date are tested last"
    (t/is (= #{"FORMULA" "_NONE" "ERROR" "STRING" "BLANK" "NUMERIC" "BOOLEAN"}
             (into #{}
                   (map (fn [e] (.name e)))
                   (CellType/values))))))

(t/deftest simple-read-write-with-strings-comparison
  (let [example {"sheet 1" [["hello" "world!"] ["foo bar"]]}]
    (t/is (= example (write-and-reread example)))))

(t/deftest simple-read-write-with-numeric-comparison
  (let [example {"sheet 1" [["hello" "world!"] ["foo bar" 1.0]]}]
    (t/is (= example (write-and-reread example)))))

(t/deftest simple-read-write-with-nil-comparison
  (let [example {"sheet 1" [[nil]]}]
    (t/is (= example (write-and-reread example)))))

(t/deftest simple-read-write-with-boolean-comparison
  (let [example {"sheet 1" [[true]]}]
    (t/is (= example (write-and-reread example)))))

(t/deftest simple-read-write-with-date-comparison
  (let [example {"sheet 1" [[(java.util.Date.)]]}]
    (t/is (= example (write-and-reread example)))))
