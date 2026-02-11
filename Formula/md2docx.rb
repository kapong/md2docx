class Md2docx < Formula
  desc "Markdown to professional DOCX converter with native Thai support"
  homepage "https://github.com/kapong/md2docx"
  version "0.1.6"

  if OS.mac?
    if Hardware::CPU.arm?
      url "https://github.com/kapong/md2docx/releases/download/v#{version}/md2docx-darwin-arm64"
      sha256 "e9d9b109a418cdea426eb4c9870e48dcc87b7f0306977e7b5c03b516b07979f5"
    else
      url "https://github.com/kapong/md2docx/releases/download/v#{version}/md2docx-darwin-x86_64"
      sha256 "0f357da775f954a8ce1aa88be8937f1cb578041bc9fffd23fac07ac09b1a559b"
    end
  elsif OS.linux?
    url "https://github.com/kapong/md2docx/releases/download/v#{version}/md2docx-linux-x86_64"
    sha256 "37eddf40259349414dd060364a0beff9506ac541a412ea198b44991d25a5cf11"
  end

  def install
    binary_name = if OS.mac?
      Hardware::CPU.arm? ? "md2docx-darwin-arm64" : "md2docx-darwin-x86_64"
    else
      "md2docx-linux-x86_64"
    end
    bin.install binary_name => "md2docx"
  end

  test do
    system "#{bin}/md2docx", "--version"
  end
end
