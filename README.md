# SCN UPLC Batch Form Maker
This program helps the laboratory professionals at Santa Cruz Nutritionals in preparing cGMP compliant forms that store all required sample preparation information. The program reads information from a password protected specification database to speed up the previously manual process of sample information retrieval. The program fills in the forms, reducing transfer errors and redundant entries - all while ensuring legibility. Any unused portions of the forms are crossed out, initialed, and date stamped. A PDF is created so that it can be reviewed for accuracy by he analyst before printing. The code, forms, images, and all other content and information are the property of Santa Cruz Nutritionals and should not be used without the written permission of Santa Cruz Nutritionals.

## Installation Guide
### Prerequisite Software
<ul>
<li>Miniconda</li>
<li>Windows</li>
</ul>

### Instructions
After downloading the repository and successfully installing miniconda on your Windows machine, navigate to the project directory. In the command prompt, enter the following: 
"conda create -n uplc_batch_helper python=3.10 -y" 

Next, enter the following:
"conda activate uplc_batch_helper"

Now, enter the following (make sure you are in the directory where libversions.txt is located):
"pip install -r libversions.txt"

If the previous command did not return any errors, you have successfully installed python and the required python packages to run the program. Now just clean up by entering the following command:
"conda deactivate"

That's it, you're done!

## Running the program
To run the program, open the command prompt. You can navigate to the project folder by entering the following:
"cd <PATH TO PROJECT FOLDER>"

Now that you are in the project folder, enter the following command:
"conda activate uplc_batch_helper && python ./scn_uplc_batch_helper_v1_0_0.py"

That's all there is to it.
