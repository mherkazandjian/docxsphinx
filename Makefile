.PHONY: nothing

nothing:

testenv:
	# .. todo:: use pipenv to deploy the installation since sphinx-build expects the plugin to be installed
	# .. todo:: and it is better not to pollute the actual dev environment

tests: clean_tests
	python setup.py install --force
	pytest -v tests --tb=no

clean_tests:
	@rm -fvr examples/**/build*
