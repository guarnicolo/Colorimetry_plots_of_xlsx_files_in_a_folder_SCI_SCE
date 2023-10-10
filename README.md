# Colorimetry_plots_of_xlsx_files_in_a_folder_SCI_SCE
This script plots all colorimetri mesuremnets from all xlsx files in a folder (both SCI and SCE).

For a specific region, a series of coating have been tested. Colorimetry mesuremnts would like to test changes in color due to the application of the coatings.
Name of the folder dhould be the same of the region.
In eache region folder there should be SCI and SCE data.
The xlsx file should contain two sheet, one pre and other post treatment.
Each sheet contains averages ad std dev of L* a* b* CIELAB coordinates.

The script plots graph with different combination of coordinates in the axis: (a* b*) and L*, (a* L*) and b*, (L* b*) and a*. Then it plots the same graph with a fixed scale on x and y axis in order to have an immediated comparison of the color changes between regions.

Then it add in the same plot the SCI and SCE data in order to visualize differences.

![SCI_SCE_OM pre e](https://github.com/guarnicolo/Colorimetry_plots_of_xlsx_files_in_a_folder_SCI_SCE/assets/81157704/eb46e8ba-24ea-41c2-9d21-34d685262bd9)
