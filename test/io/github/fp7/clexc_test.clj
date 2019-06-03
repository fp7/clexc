(ns io.github.fp7.clexc-test
  (:require [clojure.test :as t]
            [io.github.fp7.clexc :as clexc]))


(defn ^:private write-and-reread
  [example]
  (let [out-stream (java.io.ByteArrayOutputStream.)
        out (clexc/write-xlsx out-stream example)
        read (clexc/read-xlsx (java.io.ByteArrayInputStream. (.toByteArray out-stream)))]
    read))


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

(t/deftest simple-read-write-with-boolean-comparison
  (let [example {"sheet 1" [[(java.util.Date.)]]}]
    (t/is (= example (write-and-reread example)))))
