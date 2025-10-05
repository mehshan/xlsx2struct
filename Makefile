
report: test
	go tool cover -html coverage.out -o coverage.html

test:
	go test -v -coverprofile coverage.out

clean:
	rm -f coverage.out coverage.html
