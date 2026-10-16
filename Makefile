# disassemble_excel の静的解析とテスト。手元と CI (.github/workflows/ci.yml) は同じ `make lint` / `make test` を通る。
# golangci-lint はインストール不要で、go run が版を固定してビルド・起動する。
GOLANGCI_LINT_VERSION := v2.5.0

# go.work / GOFLAGS で結果を変えない (go.work があると replace の抜けを埋めて素通りし、-mod=mod だと go.sum を書き換える)
export GOWORK := off
export GOFLAGS := -mod=readonly

.PHONY: build lint test

build:
	go build -o disassemble_excel .

# go mod tidy -diff: go.mod / go.sum が tidy の形か。go list -m all: module グラフ全体が解決できるか
# (go build / go test は import が届かない module の抜けでは落ちない)
lint:
	go mod tidy -diff
	go list -m all >/dev/null
	go run github.com/golangci/golangci-lint/v2/cmd/golangci-lint@$(GOLANGCI_LINT_VERSION) run ./...

test:
	go test -race ./...
