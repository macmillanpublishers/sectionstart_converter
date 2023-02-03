# rsuite_validate container

## Dependencies
The docker container should take care of most of this, just make sure the template submodule is initialized as described [here](https://github.com/macmillanpublishers/sectionstart_converter#initialize-submodule).

## Setup container
* To build container, cd to this dir and run:
  `docker-compose build`

## Running the tool(s)
\*Note: If you need to run any of these cmds from a dir besides ./docker_rsv, add flag '-f' and the path to docker-compose file to your command, like: -f /path/to/docker-compose.yml

### run unit tests
* Unit tests run by default with cmd:
  `docker-compose up`

### run files through the tool(s)

#### running rsv _local_
* To do a *local* run of rsuite_validate on a file (without post-api's or email alerts):
	1. put a .docx file in a new folder in ./docs folder (example: ./docs/_testdir/testfile.docx_)
	2. run: `docker compose run --rm rsv python ./xml_docx_stylechecks/rsuitevalidate_main.py "/mnt/docs/testdir/testfile.docx" direct local`
	3. You will see output in the terminal as well as style_report and testfile_converted.docx being created in your _testdir_

#### running rsv
* similar to above but this would be the version for production, with emails and post-api's
	1. Update .conf/ files with true values (_smtp.txt_ and _camelPOST_urls.json_) with real values for smtp server and api destination-urls. If running in prod env. delete 'staging.txt'
	2. put a .docx file in a new folder in ./docs folder (example: ./docs/_testdir/testfile.docx_)
	3. run: `docker compose run --rm rsv python ./xml_docx_stylechecks/rsuitevalidate_main.py "/mnt/docs/testdir/testfile.docx" direct user@email.com "User Display Name"`
	4. You will see output in the terminal as well as style_report and testfile_converted.docx being created in your _testdir_ (and hopefully emails and api run too)

#### running isbncheck
* to run validator_isbncheck on a file:
	1. put a .docx file in a new folder in ./docs folder (example: ./docs/_testdir/testfile.docx_)
	2. run: `docker compose run --rm rsv python ./xml_docx_stylechecks/validator_isbncheck.py "/mnt/docs/testdir/testfile.docx"`
	3. You will see output in the terminal as well as style_report being created in your _testdir_

### run Transform/Integration Tests

* run rs_validate transform tests (these do detailed diffs on output against "validated" output):
	1. Put some test .docx files in ./test/docs
	2. If you don't have validated files, they will be generated on first run. If you have them from another environment you can drop them in "./test/validated_output"
	3. run `docker compose run --rm rsv python ./test/rsvalidate_transform_tests.py`
	(for more options re: running these tests look at this repo's test/README.md)
	4. diffs will be available in ./test/tmp

* run isbncheck transform checks: all instructions from rsv transform tests above apply, except validated files go in: "./test/vi_validated_output", and the command to run has a different executable:
	`docker compose run --rm rsv python ./test/isbncheck_transform_tests.py`
