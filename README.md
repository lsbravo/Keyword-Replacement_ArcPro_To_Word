# Keyword Replacement (ArcPro To Word)

This script is for situations where you need to run multiple intersect queries on a single point to populate a report in Word. For example, if you regularly find yourself needing to fill out values like zip codes, counties, or districts on a specific location. The tool allows you to add as many layers for intersect queries as you need but will only return one result per layer, which is why it's best for situations without overlapping polygons within the same layer.

# What You Need to Get Started

1. A Word document populated with keywords to be replaced by your intersect results. Keywords need to be unique and cannot contain spaces or semicolons. A recommended format is <<county>> for a county name lookup.
2. A shapefile with your single point location.
3. Shapefiles on the polygons to be intersected (states, regions, counties). You will also need to know how the attribute table is formatted to map the appropriate value to the keyword in Word.

# How It Works

The tool will prompt you for the various inputs. Make sure the Document Filepath includes the docx extension. The LayersToCheck portion can accept several layers. These must be accompanied by the name of the attribute that will be populated in Word and the keyword that will be replaced.

![Set Up the Parameters](https://github.com/lsbravo/Keyword-Replacement_ArcPro_To_Word/assets/121823541/675b6a65-94c7-4016-bebe-89b582d4924a)

When the tool is run, you'll see a new Word document in the same location as your original document. It will have "UPDATE" appended to the name. Open it, and you should see the keywords now mapped to the results of the intersection queries.

![Result](https://github.com/lsbravo/Keyword-Replacement_ArcPro_To_Word/assets/121823541/63dfaf35-81f2-483a-bcd4-74ca1c1e4cb5)

# How To Add It To A Custom Toolbox

In the Catalog panel, right click Toolboxes, add a new one, and rename it.

![Add_Toolbox](https://github.com/lsbravo/Keyword-Replacement_ArcPro_To_Word/assets/121823541/df9e78ea-62d0-4606-858b-ecef8e5e2d43)

Then right click the new tookbox, under New, click Script.

![Add_Script](https://github.com/lsbravo/Keyword-Replacement_ArcPro_To_Word/assets/121823541/b7c96683-452e-47b8-9b14-34befca789d6)

It'll prompt you for a name and label.

![General](https://github.com/lsbravo/Keyword-Replacement_ArcPro_To_Word/assets/121823541/890c4dbc-f845-4650-8394-4b2d11291c90)

Once you're done with that, move to the parameters tab and set these three parameters up exactly as displayed in the screenshot.

![Parameters](https://github.com/lsbravo/Keyword-Replacement_ArcPro_To_Word/assets/121823541/55cf35db-b314-40d1-8b42-706af98170e9)

Last, you'll go to the Execution tab, paste the Python script that's in this repository, and then hit OK to finish.

![Execution](https://github.com/lsbravo/Keyword-Replacement_ArcPro_To_Word/assets/121823541/9f76fa87-6a38-487c-b444-cf06b920565e)
