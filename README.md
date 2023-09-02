# Keyword Replacement (ArcPro To Word)

This script is for situations where you need to run multiple intersect queries on a single point to populate a report in Word. For example, if you regularly find yourself needing to fill out values like zip codes, counties, or districts on a specific location. The tool allows you to add as many layers for intersect queries as you need but will only return one result per feature class, which is why it's best for situations without overlapping polygons within the same feature class.

Tested in ArcGIS Pro 3.1.0.

# What You Need to Get Started

1. A Word document populated with keywords to be replaced by your intersect results. Keywords need to be unique and cannot contain spaces or semicolons. One way to guarantee no accidental replacements are made is to surround your unique keyword with <> or XXX's. See screenshot below for an example.
2. A shapefile with your single point location.
3. Shapefiles on the polygons to be intersected (states, regions, counties). You will also need to know how the attribute table is formatted to map the appropriate value to the keyword in Word.

# How It Works

All parameters are required. Make sure the Document Filepath includes the docx extension. The LayersToCheck portion can accept several layers. These must be accompanied by the name of the attribute that will be populated in Word and the keyword that will be replaced.

![Set Up the Parameters](https://github.com/lsbravo/Keyword-Replacement_ArcPro_To_Word/assets/121823541/5de0dfb8-1ff0-40ff-9a68-85eb6c9d0961)


When the tool is run, you'll see a new Word document in the same location as your original document. It will have "UPDATE" appended to the name. Open it, and you should see the keywords now mapped to the results of the intersection queries.

![Result](https://github.com/lsbravo/Keyword-Replacement_ArcPro_To_Word/assets/121823541/5e33ae1f-01b9-42a5-86aa-3681f1802693)

# Before Getting Started

Because the script interacts with Word Docs, you'll need the docx module. If you haven't already, you can open a Jupyter Notbook and import the module before getting started.

# How To Add It To A Custom Toolbox

In the Catalog panel, right click Toolboxes, add a new one, and rename it.

![Add_Toolbox](https://github.com/lsbravo/Keyword-Replacement_ArcPro_To_Word/assets/121823541/dfc5580c-ff34-4ab8-b437-9cac8d636cd9)

Then right click the new tookbox, under New, click Script.

![Add_Script](https://github.com/lsbravo/Keyword-Replacement_ArcPro_To_Word/assets/121823541/3d7a7520-d417-400f-b0eb-d1cce1201749)

It'll prompt you for a name and label.

![General](https://github.com/lsbravo/Keyword-Replacement_ArcPro_To_Word/assets/121823541/d776d909-aeaa-40e4-bad1-2f0aac989c37)

Once you're done with that, move to the parameters tab and set these three parameters up exactly as displayed in the screenshot.

![Parameters](https://github.com/lsbravo/Keyword-Replacement_ArcPro_To_Word/assets/121823541/30fac1e6-dceb-410f-a164-2d8419e3a3dd)

Last, you'll go to the Execution tab, paste the Python script that's in this repository, and then hit OK to finish.

![Execution](https://github.com/lsbravo/Keyword-Replacement_ArcPro_To_Word/assets/121823541/d36c8718-4396-42de-9cf9-131da6691224)

# Other Notes

Made with assistance from Chat GPT. Planned update to include functionality to generate multiple reports from a set of points, rather than just one.
