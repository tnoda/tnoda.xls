# tnoda.xls

A Clojure library designed to read Excel sheets (.xls). It only
supports BIFF8 format (from Excel versions 97/2000/XP/2003).

+ https://clojars.org/org.clojars.tnoda/tnoda.xls
+ https://github.com/tnoda/tnoda.xls


## Dependency Information

Leiningen dependency information:

    [org.clojars.tnoda/tnoda.xls "0.1.0"]


## Example

This library provides only one function, called `read-sheet`.

    user> (require '[tnoda.xls :as xls] '[clojure.java.io :as io])
    
    user> (-> "path/to/Tax_Year_2007_County_Income_Data.xls" io/input-stream xls/read-sheet)
    (("State Code" "County Code" "State Abbreviation" "County Name" "Total Number of Tax Returns" "Total Number of Exemptions" "Adjusted Gross Income (In Thousands)" "Wages and Salaries Incomes (In Thousands)" "Dividend Incomes (In Thousands)" "Interest Income (In Thousands)")
     (1.0 0.0 "AL" "ALABAMA" 2069212.0 4309954.0 9.2162773E7 6.7388467E7 1704870.0 2628763.0)
     (1.0 1.0 "AL" "Autauga County" 22651.0 50304.0 1052483.0 824809.0 10173.0 19682.0)
     (1.0 3.0 "AL" "Baldwin County" 75255.0 155562.0 3913872.0 2551398.0 99791.0 154580.0)
     (1.0 5.0 "AL" "Barbour County" 11644.0 23463.0 382539.0 270389.0 7678.0 14525.0)
     (1.0 7.0 "AL" "Bibb County" 9379.0 19852.0 325458.0 263779.0 2448.0 5925.0)
     (1.0 9.0 "AL" "Blount County" 23178.0 51760.0 941507.0 729702.0 9033.0 23223.0)
     (1.0 11.0 "AL" "Bullock County" 4147.0 8286.0 108430.0 82966.0 1842.0 2809.0)
     (1.0 13.0 "AL" "Butler County" 9799.0 19803.0 298185.0 218798.0 4730.0 8393.0)
     (1.0 15.0 "AL" "Calhoun County" 53376.0 109111.0 2040965.0 1518074.0 33616.0 54502.0)
     ...)

Empty cells are replaced with `nil`.

    user> (-> "path/to/japanese-demographics.xls" io/input-stream xls/read-sheet)
    (("第１表   男女別人口 (各年10月1日現在) - 総人口,日本人人口（平成12年～22年）" nil nil nil nil nil nil nil nil)
     (nil nil nil nil nil nil nil nil nil)
     ("Table 1.    Population by Sex (as of October 1 of Each Year) - Total population , Japanese population (from 2000 to 2010)" nil nil nil nil nil nil nil nil)
     (nil nil nil nil nil nil nil nil nil)
     ("(単位　千人)" nil nil nil nil nil nil nil "(Thousand persons)")
     (nil nil nil nil nil nil nil nil nil)
     (nil nil nil "総　人　口        Total population" nil nil "日本人人口       Japanese population" nil nil)
     ("　年　　次　　" nil nil "男女計" "男" "女" "男女計" "男" "女")
     ("Year " nil nil "Both sexes" "Male" "Female" "Both sexes" "Male" "Female")
     (nil nil nil nil nil nil nil nil nil)
     ("平成12年" 2000.0 "1)" 126926.0 62111.0 64815.0 125613.0 61488.0 64125.0)
     ("　　13年" 2001.0 nil 127316.0 62265.0 65051.0 125930.0 61615.0 64316.0)
     ("　　14年" 2002.0 nil 127486.0 62295.0 65190.0 126053.0 61629.0 64424.0)
     ("　　15年" 2003.0 nil 127694.0 62368.0 65326.0 126206.0 61677.0 64529.0)
     ("　　16年" 2004.0 nil 127787.0 62380.0 65407.0 126266.0 61674.0 64592.0)
     ("　　17年" 2005.0 "1)" 127768.0 62349.0 65419.0 126205.0 61618.0 64587.0)
     ("　　18年" 2006.0 nil 127901.0 62387.0 65514.0 126286.0 61630.0 64656.0)
     ("　　19年" 2007.0 nil 128033.0 62424.0 65608.0 126347.0 61635.0 64712.0)
     ("　　20年" 2008.0 nil 128084.0 62422.0 65662.0 126340.0 61609.0 64730.0)
     ("　　21年" 2009.0 nil 128032.0 62358.0 65674.0 126343.0 61586.0 64757.0) ...)


## License

Copyright © 2013 Takahiro Noda

Distributed under the Eclipse Public License, the same as Clojure.
