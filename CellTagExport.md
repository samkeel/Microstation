A quick MVBA file thrown together to get a count of all cells that contained a particular tagset in a directory for a quantity take off. 
In the microstation IDE add references to 'Microsoft scripting runtime' and 'Microsoft excel 14.0 object library'
This program needs to be run from microstation and scans the active directory looking for DGNs that contain cells with the designated cell library and exports the cell tag values to an excel file which is also saved into the active directory. 
This program cant be run from excel due to a well documented in bug between microstation and excel on reading tag data after windows xp.
