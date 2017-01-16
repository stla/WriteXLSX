{-# LANGUAGE OverloadedStrings #-}
{-# LANGUAGE TemplateHaskell #-} -- to use makeLenses

module WriteXLSX.DataframeToSheet (
    dfToCells
    , dfToSheet
    , dfToCellsWithComments
    , dfToSheetWithComments
    ) where

import WriteXLSX.Empty
-- import Codec.Xlsx (CellValue, Cell, Worksheet, Comment, 
--                   CellValue(CellBool), CellValue(CellDouble), CellValue(CellText))
import Codec.Xlsx.Types
import Control.Lens
-- import qualified Data.ByteString.Lazy as L
import qualified Data.Map.Lazy as DML
import Data.Text (Text)
import qualified Data.Text as T
import Data.Maybe (fromJust)
import Data.Aeson (decode)
import Data.Aeson.Types (Value, Object, Array, Value(Number), Value(String), Value(Bool), Value(Array), Value(Null))
import qualified Data.Aeson.Types as DAT 
import Data.HashMap.Lazy (keys)
import qualified Data.HashMap.Lazy as DHL
import Data.ByteString.Lazy.Internal (packChars)
import qualified Data.ByteString.Lazy.Internal as DBLI
import qualified Data.Vector as DV
import Text.Regex 

makeLenses ''Comment


extractKeys :: Regex -> String -> [String]
extractKeys reg s = 
  case matchRegexAll reg s of
    Nothing -> []
    Just (_, _, after, matched) -> matched ++ extractKeys reg after

-- problem: decode does not preserve order of keys
-- c'est à cause de la minuscule de include on dirait


dfToColumns :: String -> ([Array], [Text])
dfToColumns df = (map (\key -> valueToArray $ fromJust $ DHL.lookup key dfObject) colnames, colnames)
                   where dfObject = fromJust $ (decode $ packChars df :: Maybe Object)
                         colnames = T.pack <$> extractKeys (mkRegex "\"([^:|^\\,]+)\":") df

valueToArray :: Value -> Array
valueToArray value = x
                     where Array x = value

columnToExcelColumn :: Array -> [Maybe CellValue]
columnToExcelColumn column = DV.toList $ DV.map valueToCellValue column


valueToCellValue :: Value -> Maybe CellValue 
valueToCellValue value = 
    case value of
        (Number x) -> Just (CellDouble (realToFrac x :: Double))
        (String x) -> Just (CellText x)
        (Bool x) -> Just (CellBool x)
        Null -> Nothing

commentsToExcelComments :: Array -> String -> [Maybe Comment]
commentsToExcelComments column author = DV.toList $ 
                                          DV.map (\x -> valueToComment x author) column

valueToComment :: Value -> String -> Maybe Comment 
valueToComment value author = 
    case value of
        (String x) -> Just $ set commentAuthor (T.pack author) $ 
                               set commentText (XlsxText $ x) emptyComment
        Null -> Nothing


-- j :: Int -- column index
-- j = 1
-- excelcol = columnToExcelColumn $ (dfToColumns df colnames) !! j
-- x = map (\(i, maybeCell) -> ((i,j), set cellValue maybeCell emptyCell)) $
--   zip [1..length excelcol] excelcol
-- map sur j et concat 

-- je rajoute ncols pour ColumnWidths
-- les widths sont peut-être auto pour les nombres et les dates (comme Excel 2003)

dfToCells :: String -> Bool -> (CellMap, Int)
dfToCells df header = (DML.fromList $ concat $ map f [1..ncols], ncols)
      where f j = map (\(i, maybeCell) -> ((i,j), set cellValue maybeCell emptyCell)) $ 
                        zip [1..length excelcol] excelcol
                    where excelcol0 = columnToExcelColumn $ dfCols !! (j-1)
                          excelcol = if header 
                                        then (Just (CellText (colnames !! (j-1)))) : excelcol0
                                        else excelcol0
            (dfCols, colnames) = dfToColumns df
            ncols = length dfCols

-- résultat de widths : Excel répare...
--  je crois qu'il faut rajouter un "Fill" dans _styleSheetFills du styleSheet ! :-(
-- dans Book1Comments j'ai une width et :
-- <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/>
-- pour en être sur : write un fichier, unzip, repair with Excel, unzip, compare !
-- il faudrait une fonction qui lit le ColumnsWidth dans le ws et qui crée le StyleSheet
-- (de même pour les CellXfs)
-- j'ai mis le Fill et Excel répare encore !
-- j'ai jeté un oeil, Excel ajoute un borderdiagonal dans l'unique Border
-- essaye _borderDiagonal = Just (BorderStyle {_borderStyleColor = Nothing, _borderStyleLine = Nothing})

dfToSheet :: String -> Bool -> Worksheet
dfToSheet df header = set wsColumns [widths] $ set wsCells cells emptyWorksheet
                        where (cells, ncols) = dfToCells df header
                              widths = ColumnsWidth {cwMin = 1,
                                                     cwMax = ncols,
                                                     cwWidth = 30,
                                                     cwStyle = 1}

dfToCellsWithComments :: String -> Bool -> String -> String -> CellMap
dfToCellsWithComments df header comments author = DML.fromList $ concat $ map f [1..length dfCols]
      where f j = map (\(i, maybeCell, maybeComment) -> 
                         ((i,j), set cellComment maybeComment $ 
                                   set cellValue maybeCell emptyCell)) $ 
                        zip3 [1..length excelcol] excelcol excelComments
                    where excelcol0 = columnToExcelColumn $ dfCols !! (j-1)
                          excelcol = if header 
                                        then (Just (CellText (colnames !! (j-1)))) : excelcol0
                                        else excelcol0
                          excelComments0 = commentsToExcelComments (dfComments !! (j-1)) author
                          excelComments = if header
                                             then Nothing : excelComments0
                                             else excelComments0 
            (dfCols, colnames) = dfToColumns df
            (dfComments, _) = dfToColumns comments


dfToSheetWithComments :: String -> Bool -> String -> String -> Worksheet
dfToSheetWithComments df header comments author = 
  set wsCells (dfToCellsWithComments df header comments author) emptyWorksheet
