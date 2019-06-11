; Copyright (c) 2019 Finn Petersen
;
; This program and the accompanying materials are made
; available under the terms of the Eclipse Public License 2.0
; which is available at https://www.eclipse.org/legal/epl-2.0/
;
; SPDX-License-Identifier: EPL-2.0

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

(t/deftest simple-read-write-cycles
  (let [test-date (java.util.Date.)]
    (t/are [example] (= example (write-and-reread example))
      {"sheet 1" [["hello" "world!"] ["foo bar"]]}
      {"sheet 1" [["hello" "world!"] ["foo bar" 1.0]]}
      {"sheet 1" [[nil]]}
      {"sheet 1" [[true]]}
      {"sheet 1" [[]]})))


(t/deftest simple-read-write-cycle-should-drop-wrapper
  (let [test-date (java.util.Date.)]
    (t/are [expected sheet] (= expected (write-and-reread sheet))
      {"sheet 1" [["hello" "world!"] ["foo bar"]]}
      {"sheet 1" [[{:value "hello"} "world!"] ["foo bar"]]}

      {"sheet 1" [["hello" "world!"] ["foo bar" 1.0]]}
      {"sheet 1" [["hello" "world!"] ["foo bar" {:value 1.0}]]}

      {"sheet 1" [[nil]]}
      {"sheet 1" [[{}]]}

      {"sheet 1" [[nil]]}
      {"sheet 1" [[{:value nil}]]}

      {"sheet 1" [[true]]}
      {"sheet 1" [[{:value true}]]}

      {"sheet 1" [[test-date]]}
      {"sheet 1" [[{:value test-date}]]})))

(t/deftest default-format-should-not-reside-in-meta
  (t/is (= nil
           (meta
            (-> (write-and-reread {"sheet 1" [[(java.util.Date.)]]})
                (get "sheet 1")
                (ffirst))))))

(t/deftest custom-format-should-reside-in-meta
  (t/is (= {:cell-format "m/d/yy"}
           (meta
            (-> (write-and-reread {"sheet 1" [[(with-meta
                                                 {:value (java.util.Date.)}
                                                 {:cell-format "m/d/yy"})]]})
                (get "sheet 1")
                (ffirst))))))

(t/deftest non-builtin-format-should-reside-in-meta
  (t/is (= {:cell-format "h"}
           (meta
            (-> (write-and-reread {"sheet 1" [[(with-meta
                                                 {:value (java.util.Date.)}
                                                 {:cell-format "h"})]]})
                (get "sheet 1")
                (ffirst))))))

(t/deftest ints-should-be-formattable
  (t/is (= {:cell-format "00#,###"}
           (meta
            (-> (write-and-reread {"sheet 1" [[(with-meta
                                                 {:value 400000}
                                                 {:cell-format "00#,###"})]]})
                (get "sheet 1")
                (ffirst))))))
