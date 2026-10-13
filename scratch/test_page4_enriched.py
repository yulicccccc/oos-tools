import fitz

doc = fitz.open(r"C:\Users\qchen\OneDrive - Professional Compounding Centers of America, Inc\Documents\OOS\ScanRDI OOS P2 template.pdf")
page4 = doc[3]

p24 = (
    "PHASE II RETEST UNDER SUBMISSION ETX-260914-0470:\r"
    "In accordance with Phase II investigation procedures under Form 3.100.019.F02, additional sample vials from lot 2608-216 were submitted "
    "for retesting under submission ETX-260914-0470.\r \r"
    "The retest sample was prepared by Andrew Carrillo on 14 Sep 2026 and processed by Ayomide Odugbesi on 15 Sep 2026 in accordance with "
    "MICRO-SOP-5 (USP <71> / EP 2.6.1 Sterility Test) via membrane filtration (60 gm filtered across 4 vials). The retest was conducted in certified "
    "ISO 5 BSC E001314 within Cleanroom Suite 115 following standard multi-barrier disinfection and strict aseptic technique. All culture media "
    "(FTM Lot 2567320, TSB Lot 2567280) and reagents (Fluid D Lot 687774, Bacteriostatic Water Lot NM7681) were within valid expiration dates. "
    "Media bottles were incubated in designated incubators E001356 and E001357. Negative controls were handled concurrently and remained sterile.\r \r"
    "On 21 Sep 2026 (Day 06 of incubation), microbial growth was observed in 1 x 100mL TSB media bottle by analyst Andrew Carrillo and confirmed "
    "by supervisor Robin Seymour. The positive bottle was submitted for Microbial Identification under ETX-260921-0498. Confirmatory identification "
    "definitively identified the isolate as Microbacterium sp. PM5 (Gram-positive short rods), matching the identical species and strain designation (PM5), "
    "morphology (white shiny lawn growth), media selection (TSB), and incubation kinetics (Day 6) of the initial failure under ETX-260902-0505 "
    "(ETX-260910-0290). All other samples processed in the batch on 15 Sep 2026 tested negative.\r \r"
    "Upon analyzing environmental monitoring results for the retest processing on 15 Sep 2026, no microbial growth was detected on operator personnel "
    "fingertip touch plates, settling plates, or ISO 5 BSC surface contact plates on the date of testing (15 Sep 2026), the preceding testing date (14 Sep 2026), "
    "or the subsequent testing date (16 Sep 2026). Weekly active surface monitoring of Cleanroom Suite 115 showed no growth; weekly active air sampling in "
    "the outermost ISO 8 cleanroom 115 recovered 9 CFUs (ETX-260923-0402), identified as Kocuria indica, Brevibacterium sp. CS2, Corynebacterium sp, Micrococcus luteus, "
    "Paracoccus yeei, Staphylococcus hominis, and Kocuria rhizophila. 100% of these environmental isolates differed distinctly in genus and morphology from the product "
    "contamination isolate (Microbacterium sp. PM5). The ISO 5 critical zones in BSC E001314 remained completely sterile (0 CFU) and in optimal control throughout testing, "
    "with positive pressure cascades preventing anteroom air ingress. (Refer to Table 1 for sample details, Table 2 for environmental monitoring from processing "
    "performed on 04 Sep 2026, and Table 3 for environmental monitoring from processing performed on 15 Sep 2026).\r \r"
    "EVALUATION OF RETEST FINDINGS & JUSTIFICATION OF PRIOR SCANRDI RESULT:\r"
    "The successful replication of microbial recovery in the independent retest—conducted by a different processing analyst (Ayomide Odugbesi) "
    "in a different cleanroom suite (Cleanroom Suite 115 vs. Suite 114) and a different prepping/reading analyst (Andrew Carrillo)—conclusively "
    "refutes the preliminary hypothesis of analyst handling error during reconstitution. If the initial recovery had been an isolated artifact "
    "of analyst technique, the identical microorganism (Microbacterium sp. PM5) would not have been recovered on the identical incubation day in the retest.\r \r"
    "Importantly, the investigation evaluated why the prior rapid ScanRDI sterility test conducted on this lot (submission ETX-260807-0602, test record 081326-2017-2) "
    "yielded a passing result (Count 14/0, Pass). First, per EagleTrax records for ETX-260807-0602, only 1 single container (Number of Items in Batch: 1, 30 mL/gm) "
    "was submitted and tested via ScanRDI, whereas USP <71> compendial testing analyzed 3 vials (45 gm filtered per media) for initial test ETX-260902-0505 and 4 vials "
    "for retest ETX-260914-0470. Second, official records confirm that 'No Method Suitability on file' exists for ScanRDI with this specific peptide analyte formulation "
    "(Cagrilinitide / Retatrutide). Third, in compounded parenteral formulations, microbial contamination at low bioburden levels is well recognized to exhibit a non-uniform, "
    "heterogeneous distribution across individual batch containers. Under single-vial sampling (n = 1), random chance naturally selects a container devoid of viable cells—as "
    "occurred with the single vial tested under ETX-260807-0602—whereas multi-container sampling captures vials harboring viable cells. When viable cells were introduced into "
    "14-day compendial broth enrichment (USP <71> membrane filtration), the organisms were successfully enriched and recovered in TSB on Day 6 in both independent testing events "
    "(ETX-260902-0505 and ETX-260914-0470), yielding the identical strain (Microbacterium sp. PM5). Therefore, the passing ScanRDI result and the two failing USP <71> compendial "
    "results are scientifically consistent and attributable to non-uniform distribution of low-concentration microbial bioburden within the compounded lot. Consequently, "
    "inherent product contamination in Lot 2608-216 is confirmed, the initial failing result for ETX-260902-0505 is valid, retest ETX-260914-0470 is confirmed failing, "
    "and Lot 2608-216 fails sterility requirements."
)

for w in page4.widgets():
    if w.field_name == 'Text Field24':
        w.field_value = p24
        w.text_fontsize = 7.7
        w.text_font = 'Helv'
        w.update()

doc.save("scratch/test_page4_enriched.pdf")
print("Saved test_page4_enriched.pdf")

# Render image of Page 4 to check layout
page4_new = fitz.open("scratch/test_page4_enriched.pdf")[3]
pix = page4_new.get_pixmap(dpi=150)
pix.save("scratch/preview_p4_enriched.png")
print("Saved scratch/preview_p4_enriched.png")
