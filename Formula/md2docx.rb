class Md2docx < Formula
  desc "Markdown to professional DOCX converter with native Thai support"
  homepage "https://github.com/kapong/md2docx"
  version "0.1.4"

  if OS.mac?
    if Hardware::CPU.arm?
      url "https://github.com/kapong/md2docx/releases/download/v#{version}/md2docx-darwin-arm64"
      sha256 "7ab3d1571ab0eae0f3f65cd3117a54e58046952e71559b6ec2ef7ba36e82c060"
    else
      url "https://github.com/kapong/md2docx/releases/download/v#{version}/md2docx-darwin-x86_64"
      sha256 "c5b2cfbcc0db5835b7680de4cbc9167ad1128255c9d6303fd23f406a014fc82d"
    end
  elsif OS.linux?
    url "https://github.com/kapong/md2docx/releases/download/v#{version}/md2docx-linux-x86_64"
    sha256 "e871fe848120bdec4b37e22ec7a2b75d8229cfb659af9f195c188cce34582c22"
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
