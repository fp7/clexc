(ns io.github.fp7.clexc-test
  (:require [clojure.test :as t]
            [io.github.fp7.clexc :as clexc]))


(t/deftest simple-read-write-comparison
  (let [example {"sheet 1" [["hello" "world!"] ["foo bar"]]}
        out-stream (java.io.ByteArrayOutputStream.)
        out (clexc/write-xlsx out-stream example)
        read (clexc/read-xlsx (java.io.ByteArrayInputStream. (.toByteArray out-stream)))]
    (t/is (= example read))))
