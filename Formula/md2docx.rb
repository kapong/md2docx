class Md2docx < Formula
  desc "Markdown to professional DOCX converter with native Thai support"
  homepage "https://github.com/kapong/md2docx"
  version "0.2.1"

  if OS.mac?
    if Hardware::CPU.arm?
      url "https://github.com/kapong/md2docx/releases/download/v#{version}/md2docx-darwin-arm64"
      sha256 "a66da5b9686032e1bb57afed87ea05a561fe37598b47a96354b52e5120c67f24"
    else
      url "https://github.com/kapong/md2docx/releases/download/v#{version}/md2docx-darwin-x86_64"
      sha256 "1a40b690d68588bcd8a8bafffaea10247039f52e198c52888940a1df020aa715"
    end
  elsif OS.linux?
    url "https://github.com/kapong/md2docx/releases/download/v#{version}/md2docx-linux-x86_64"
    sha256 "bc49edec1ca943aa4bb681a866039e429413ffbf77bba514254a0a69750452f1"
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
