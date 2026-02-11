class Md2docx < Formula
  desc "Markdown to professional DOCX converter with native Thai support"
  homepage "https://github.com/kapong/md2docx"
  version "0.1.7"

  if OS.mac?
    if Hardware::CPU.arm?
      url "https://github.com/kapong/md2docx/releases/download/v#{version}/md2docx-darwin-arm64"
      sha256 "6546692c0c58223dcde36e4f5b85b540a05720b4ce20224926f3e60fe0249e89"
    else
      url "https://github.com/kapong/md2docx/releases/download/v#{version}/md2docx-darwin-x86_64"
      sha256 "04673c46070f77668a5aae766a8c090b77e4c73775cc24ffb1806ebea189c492"
    end
  elsif OS.linux?
    url "https://github.com/kapong/md2docx/releases/download/v#{version}/md2docx-linux-x86_64"
    sha256 "4590827de274ca66c3953585fd0fc82ce5dec0bc2a9f460d3d5309319a5be50c"
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
