# fyst-primecam-tolerance-calculator
A set of Excel spreadsheets containing parametric models of Prime-Cam's 3-Lens, 4-Lens, and EoR-Spec detector designs, as well as Python Code to set component tolerances of the 3-Lens model.


### Components
##### Tolerance Calculators
These are three excel spreadsheets containing a parameterized model of the 410 GHz, 850 GHz, and EoR-Spec instruments of Prime-Cam. Specifically, each component/critical axial dimension is given a row in the spreadsheet. Included in each row is the part's label on the included diagrams, temperature stage, name, inputted and ideal part length, calculated and ideal x-position, deltas for each, and inputted machine tolerances and measurement uncertainties

Within the model, the x-direction points out the front of the instrument. A part's position is defined as the distance from the attachment point between the upper shell and Prime-Cam 40K stage and the back end of the part relative to the x-direction. When the length of a part in altered by the user in comparison to its ideal value, all dependent part positions (including potentially its own) will automatically update. The deltas will also update, and conditional formatting is used to show whether a component's position and length violate the given tolerances within the given uncertainty. A second page on the excel file finds the critical dimensions used to define the lens tolerances using the calculated values and determines their respective deltas. The previously calculated lens spacing tolerances ($/pm$1mm) were used for this spreadsheet.

