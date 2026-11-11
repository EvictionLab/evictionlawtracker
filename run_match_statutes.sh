python3 match_statutes.py \
  --workbook ./EvictionLawTracker.xlsx \
  --xml-root ./xmldata \
  --output ./LSCEvictionLaws_StateTerritory_Data_matched.xlsx \
  --report ./LSCEvictionLaws_StateTerritory_Data_matched.by_year.csv \
  --index-db ./data/section_index.sqlite \
  --target-year 2021
