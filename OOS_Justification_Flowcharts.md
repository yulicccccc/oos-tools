# OOS EM Growth Justification Flowcharts by Test Method (Updated)

Based on our final, bulletproof logic, the **Core Logic Engine** applies a 4-step shielding mechanism. What changes between the 3 tests is the **Phasing** (the timeline of when the sample was exposed). The updated decision trees below map out the ultimate defense logic.

---

## 1. Celsis Decision Tree (Two-Phase Exposure)
Celsis explicitly involves **Processing** and **Aliquoting** phases. 

```mermaid
flowchart TD
    Start([Growth Detected in Celsis EM]) --> Q1{Phase of Growth?}
    
    Q1 -- Processing Phase --> Q2_Pro{Rule 1: Is Isolate Identical?}
    Q1 -- Aliquoting Phase --> Q2_Alq{Rule 1: Is Isolate Identical?}
    
    Q2_Pro -- No --> J1[Justification:<br>Isolates not identical during Processing. Independent event.] --> Q3_Pro
    Q2_Pro -- Yes --> Q3_Pro{Rule 2: Where was Growth?}
    
    Q3_Pro -- ISO 8 / Outer Room --> J2[Justification:<br>Detected in ISO 8 background, but manipulation was strictly in ISO 5.] --> Q4_Pro
    
    Q4_Pro{Rule 3: Transfer Pathway?} -- Analyst/Surfaces Clean --> J3[Justification:<br>No viable transfer pathway to ISO 5 BSC.] --> Q5_Pro
    
    Q5_Pro{Rule 4: Other Samples?} -- Clean --> J4[Justification:<br>Lack of contamination in other samples supports optimal conditions.] --> Final
    
    Q2_Alq -- No --> J5[Justification:<br>Isolates not identical during Aliquoting. Independent event.] --> Q3_Alq
    Q2_Alq -- Yes --> Q3_Alq{Rule 2: Where was Growth?}
    
    Q3_Alq -- ISO 8 / Outer Room --> J6[Justification:<br>Detected in ISO 8 background, but manipulation was strictly in ISO 5.] --> Q4_Alq
    
    Q4_Alq{Rule 3: Transfer Pathway?} -- Analyst/Surfaces Clean --> J7[Justification:<br>No viable transfer pathway to ISO 5 BSC.] --> Q5_Alq
    
    Q5_Alq{Rule 4: Other Samples?} -- Clean --> J8[Justification:<br>Lack of contamination in other samples supports optimal conditions.] --> Final
    
    Final([Final Conclusion:<br>Facility's layered disinfection and cleanroom design blocked contamination. Lab error is minimal, original result is valid.])
    
    classDef question fill:#f9f2f4,stroke:#d07b8a,stroke-width:2px,color:#a12036,font-weight:bold
    classDef justify fill:#e6f3ff,stroke:#4a90e2,stroke-width:2px,color:#003366
    classDef conclusion fill:#e6ffe6,stroke:#2ca02c,stroke-width:2px,color:#004d00,font-weight:bold
    
    class Q1,Q2_Pro,Q2_Alq,Q3_Pro,Q3_Alq,Q4_Pro,Q4_Alq,Q5_Pro,Q5_Alq question
    class J1,J2,J3,J4,J5,J6,J7,J8 justify
    class Final conclusion
```

---

## 2. Scan RDI Decision Tree (Single-Phase Exposure)
Scan RDI has a single **Processing** event. The logic heavily relies on the physical separation between ISO 8 and ISO 5.

```mermaid
flowchart TD
    Start([Growth Detected in Scan RDI EM]) --> Q1{Rule 1: Is Isolate Identical?}
    
    Q1 -- No --> J1[Justification:<br>Isolates not identical to sample. Independent event.] --> Q2
    Q1 -- Yes --> Q2{Rule 2: Where was Growth?}
    
    Q2 -- ISO 8 (Outer Room) --> J2[Justification:<br>Detected in ISO 8 background, whereas sample manipulation occurred strictly within ISO 5.] --> Q3
    Q2 -- Floor (ISO 7) --> J3[Justification:<br>Detected on ISO 7 floor, but sample manipulation occurred strictly within ISO 5.] --> Q3
    
    Q3{Rule 3: Transfer Pathway?} -- Analyst/Surfaces Clean --> J4[Justification:<br>No viable transfer pathway existed from the ISO 8/7 area to the ISO 5 BSC.] --> Q4
    
    Q4{Rule 4: Other Samples?} -- Clean --> J5[Justification:<br>Lack of contamination in other samples supports optimal conditions.] --> Final
    
    Final([Final Conclusion:<br>Facility's layered disinfection and cleanroom design blocked contamination. Lab error is minimal, original result is valid.])
    
    classDef question fill:#f9f2f4,stroke:#d07b8a,stroke-width:2px,color:#a12036,font-weight:bold
    classDef justify fill:#e6f3ff,stroke:#4a90e2,stroke-width:2px,color:#003366
    classDef conclusion fill:#e6ffe6,stroke:#2ca02c,stroke-width:2px,color:#004d00,font-weight:bold
    
    class Q1,Q2,Q3,Q4 question
    class J1,J2,J3,J4,J5 justify
    class Final conclusion
```

---

## 3. USP <71> Decision Tree (Processing + Subculture)
USP <71> introduces a **Subculture / Day 14+** phase. The justification must distinguish whether the growth happened on the "day of processing" or "during the week of subculture".

```mermaid
flowchart TD
    Start([Growth Detected in USP 71 EM]) --> Q1{Phase of Growth?}
    
    Q1 -- Day of Processing --> Q2_Pro{Rule 1: Is Isolate Identical?}
    Q1 -- Week of Subculture --> Q2_Sub{Rule 1: Is Isolate Identical?}
    
    Q2_Pro -- No --> J1[Justification:<br>Isolates not identical on the day of processing. Independent event.] --> Q3_Pro
    Q2_Pro -- Yes --> Q3_Pro{Rule 2: Where was Growth?}
    
    Q3_Pro -- ISO 8 (Outer Room) --> J2[Justification:<br>Detected in ISO 8 background, whereas sample manipulation occurred strictly within ISO 5.] --> Q4_Pro
    
    Q4_Pro{Rule 3: Transfer Pathway?} -- Analyst/Surfaces Clean --> J3[Justification:<br>No viable transfer pathway existed from the ISO 8/7 area to the ISO 5 BSC.] --> Q5_Pro
    
    Q5_Pro{Rule 4: Other Samples?} -- Clean --> J4[Justification:<br>Lack of contamination in other samples supports optimal conditions.] --> Final
    
    Q2_Sub -- No --> J5[Justification:<br>Isolates not identical during the week of subculture. Independent event.] --> Q3_Sub
    Q2_Sub -- Yes --> Q3_Sub{Rule 2: Where was Growth?}
    
    Q3_Sub -- ISO 8 (Outer Room) --> J6[Justification:<br>Detected in ISO 8 background, whereas subculture manipulation occurred strictly within ISO 5.] --> Q4_Sub
    
    Q4_Sub{Rule 3: Transfer Pathway?} -- Analyst/Surfaces Clean --> J7[Justification:<br>No viable transfer pathway existed from the ISO 8/7 area to the ISO 5 BSC.] --> Q5_Sub
    
    Q5_Sub{Rule 4: Other Samples?} -- Clean --> J8[Justification:<br>Lack of contamination in other samples supports optimal conditions.] --> Final
    
    Final([Final Conclusion:<br>Facility's layered disinfection and cleanroom design blocked contamination. Lab error is minimal, original result is valid.])
    
    classDef question fill:#f9f2f4,stroke:#d07b8a,stroke-width:2px,color:#a12036,font-weight:bold
    classDef justify fill:#e6f3ff,stroke:#4a90e2,stroke-width:2px,color:#003366
    classDef conclusion fill:#e6ffe6,stroke:#2ca02c,stroke-width:2px,color:#004d00,font-weight:bold
    
    class Q1,Q2_Pro,Q2_Sub,Q3_Pro,Q3_Sub,Q4_Pro,Q4_Sub,Q5_Pro,Q5_Sub question
    class J1,J2,J3,J4,J5,J6,J7,J8 justify
    class Final conclusion
```
