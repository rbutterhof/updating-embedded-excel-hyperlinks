# Updating the embedded links in the US Newspapers Currently Recieved list

I wanted to swap out Stacks links for handle links in the LibGuide so offsite users could get more infomation about how to access Stacks. If you click on the Stacks link offsite, you get a page not found error.  Both of the links end with the Library of Congress Control Number (LCCN) so I wanted something to swap out the prefix only on the links. 

This repo includes the original Excel file and the VBA script used to update it. 

As part of this project, I realized that some of the Stacks titles didn't have handles, so I got those created. I also had to make some small fixes as some of the title were not SER titles, and they needed a different prefix: https://hdl.loc.gov/loc.gdc/stacksnewspaper. instead of https://hdl.loc.gov/loc.sgp/npe.
