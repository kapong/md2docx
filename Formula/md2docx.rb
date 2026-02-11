class Md2docx < Formula
  desc "Markdown to professional DOCX converter with native Thai support"
  homepage "https://github.com/kapong/md2docx"
  version "0.1.4"

  if OS.mac?
    if Hardware::CPU.arm?
      url "https://github.com/kapong/md2docx/releases/download/v#{version}/md2docx-darwin-arm64"
      sha256 "18f6a313fd274e416e773f369cab6a393ecddf528b2933d00cff20f325a799d4"
    else
      url "https://github.com/kapong/md2docx/releases/download/v#{version}/md2docx-darwin-x86_64"
      sha256 "c9dfa5cc1d23336413a0d82018d2d0bb56e08eba98588ff8cdd64f2c6d559b1d"
    end
  elsif OS.linux?
    url "https://github.com/kapong/md2docx/releases/download/v#{version}/md2docx-linux-x86_64"
    sha256 "c9fe8c1f4b4c5b82bada8c1d72c52805ff0629485e6f5446a56a24c7d88fee50"
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
