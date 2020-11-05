.PHONY: nothing

nothing:

testenv:
	# .. todo:: use pipenv to deploy the installation since sphinx-build expects the plugin to be installed
	# .. todo:: and it is better not to pollute the actual dev environment

tests:
	sphinx-build -b docx source build

clean:
	@rm -fvr examples/**/build*
