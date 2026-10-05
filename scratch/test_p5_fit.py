import re

# Test the split
p11_clean = (
    "Monthly cleaning and disinfection of Cleanroom Suite 114 and its Biosafety Cabinets was performed on "
    "30 Aug 2026 as per MICRO-SOP-9, and all hydrogen peroxide chemical indicators met applicable acceptance "
    "criteria. Attached Tables 2 and 3 present the complete environmental monitoring results for the duration "
    "of testing, incubated in accordance with MICRO-SOP-2."
)

p12_em = (
    "Review of daily environmental monitoring showed zero microbial recovery across all personnel touch plates, "
    "settling plates, and Biosafety Cabinet work surfaces for processing on 01 Sep 2026 and aliquoting on 08 Sep 2026, "
    "as well as on their respective preceding and subsequent bracketing dates. Weekly monitoring identified 1 CFU in "
    "outermost ISO 8 anteroom 114 on 04 Sep 2026 under sample ETX-260914-0487, characterized as Gram-positive coccobacilli "
    "and Gram-positive rods, and 1 CFU in the same area on 10 Sep 2026 under sample ETX-260921-0520, characterized as "
    "Gram-positive cocci. No microbial recovery was observed from ISO 7 cleanrooms 114A or 114B, and weekly surface "
    "sampling across Suite 114 showed no growth across all locations."
)

p13_seg = (
    "All testing activities were performed within validated ISO 5 Biosafety Cabinets E001316 and E001798 located in "
    "cleanrooms 114B and 114A. Sample containers and media remained enclosed during transfer through anteroom 114 "
    "and were never exposed to ambient anteroom air. Because no organism was recovered from ETX-260828-0527 during "
    "subculture, a direct microbiological comparison could not be performed. However, closed material transfer, "
    "negative critical zone monitoring, and positive pressure airflow cascades confirmed that the positive result was "
    "not linked to the laboratory environment."
)

p14_batch = (
    "All interviewed analysts followed cleaning and handling procedures per MICRO-SOP-9 and MICRO-SOP-44 with no "
    "documented deviations. ETX-260828-0527 was the first sample processed by analyst Gabrielle Surber on 01 Sep 2026, "
    "and all other samples processed, incubated, aliquoted, and read within the testing batch tested negative, confirming "
    "no cross-contamination occurred."
)

p15_hist = (
    "An analysis of the six-month sample history for Optimal Balance Pharmacy under account E19193 indicates that "
    "Eagle Analytical processed [X] samples for Celsis Sterility testing with no prior occurrences of an "
    "out-of-specification or positive result for this analyte during this period."
)

p16_lot = (
    "A review of the lot history shows that there have been no additional submissions of sample lot LG342010349 for "
    "retesting using Celsis Sterility Testing or any other sterility method."
)

p17_conc = (
    "Based on the laboratory investigation, no assignable laboratory cause was identified involving the analyst, "
    "instrument, reagents, supplies, or monitored laboratory environment. The initial Celsis OOS result therefore "
    "remains valid in accordance with the applicable laboratory OOS procedure."
)

p5_text = "\n\n".join([p11_clean, p12_em, p13_seg, p14_batch, p15_hist, p16_lot, p17_conc])
print("Total words on Page 5:", len(p5_text.split()))
print("Total chars on Page 5:", len(p5_text))
