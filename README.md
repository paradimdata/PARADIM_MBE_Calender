# PARADIM Cornell Calendar
This is a script written to help track time instruments were in use for specific proposals. This script scrapes all events from the Cornell lab Google calendar kept by Matt Barone. Events are then sorted and only events with proposal numbers are added to an excel sheet along with the times when the proposal was using an instrument. The current working version is main.py. 

## Usage
#### Installation
First, set up a virtual environment. This script was written and tested in a conda environment. 

	conda create -n googlecalendarapi python -y
	conda activate googlecalendarapi
Then install required packages in your new environment

	pip install pandas
	pip install datetime
	pip install os

### Google QuickStart Instructions
Next, follow the google quickstart instructions. Before starting the the Quickstart instructions, you must make sure one of your google accounts has been added to the application as a user. The google quickstart instructions can be found [here](https://developers.google.com/calendar/api/quickstart/python). Enabling the API and configuring the OAuth screen should be very simple and should just require clicking through a couple of screens. For the credentials step, make sure you download the file and put it in the same folder where the main.py script is and rename it to credentials.json. Finally, make sure you execute the following pip install command in your environment.

	pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
After following these steps run the main.py script. The first time running the script should take you to a browser where you are asked to sign in to recieve your token.json file. This may take more than one attempt to execute properly. When you recieve your token.json file make sure it is in the same folder as the main.py file. 	
