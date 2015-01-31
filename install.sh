sudo apt-get install python2.7 python-dev libxml2-dev libxslt-dev zlib1g-dev

# the user is inside a virtual environment
if [ -n "$VIRTUAL_ENV" ]; then

	echo "Using a virtual environment: just run pip install"

	pip install -r requirements.txt

else
	echo "Not using a virtual environment: install setuptools and pip system-wide"

	sudo apt-get install python-setuptools

	sudo easy_install pip

	sudo pip install -r requirements.txt
fi
