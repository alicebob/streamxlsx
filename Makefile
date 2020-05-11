test:
	go test

example:
	go test -run Example

unzip:
	go test -run Example && rm -rf ./ex/ && mkdir ex && (cd ex && unzip ../example.xlsx)

bench:
	go test -bench .
