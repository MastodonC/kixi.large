(ns kixi.large
  (:require [kixi.large.legacy :as ll]
            [tablecloth.api :as tc])
  (:import org.apache.poi.ss.usermodel.Workbook
           (org.apache.poi.xssf.usermodel XSSFWorkbook)
           (org.apache.poi.ss.usermodel Workbook Sheet Cell Row CellType
                                        Row$MissingCellPolicy
                                        HorizontalAlignment
                                        VerticalAlignment
                                        BorderStyle
                                        FillPatternType
                                        FormulaError
                                        WorkbookFactory DateUtil
                                        IndexedColors CellStyle Font
                                        CellValue Drawing CreationHelper)))

(defn excel-tab-string [tab-name]
  (if (< 31 (count tab-name))
    (subs tab-name 0 31)
    tab-name))

(defn add-image! [^Sheet sheet
                  {::keys [image col-anchor row-anchor]
                   :or {col-anchor 10
                        row-anchor 1}
                   :as config}]
  (try
    (when image
      (let [workbook (.getWorkbook sheet)
            pic-idx (.addPicture workbook image Workbook/PICTURE_TYPE_PNG)
            helper (.getCreationHelper workbook)
            drawing (.createDrawingPatriarch sheet)
            anchor (.createClientAnchor helper)
            col-width 13.0
            row-height 36.0
            _ (.setCol1 anchor col-anchor)
            _ (.setRow1 anchor row-anchor)
            ;; Picture pict = drawing.createPicture(anchor, pictureIdx);
            ;; pict.resize();
            pict (.createPicture drawing anchor pic-idx)]
        (.resize pict col-width row-height)))
    (catch Exception e
      (throw (ex-info "Failed to add image" config e)))))

(defn add-sheet! [^Workbook workbook {::keys [sheet-name data images]}]
  (let [sheet (.createSheet workbook sheet-name)
        rows (into [(into [] (map name) (tc/column-names data))]
                   (map vec)
                   (tc/rows data :as-seqs))]
    (ll/add-rows! sheet rows)
    (when (seq images)
      (doseq [i images]
        (add-image! sheet i)))
    workbook))


(defn create-workbook [sheet-specs]
  (let [workbook (XSSFWorkbook.)]
    (doseq [s sheet-specs]
      (add-sheet! workbook s))
    workbook))

(defn save-workbook! [wb file-name]
  (ll/save-workbook! file-name wb))
