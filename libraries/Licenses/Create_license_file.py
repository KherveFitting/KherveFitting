# create_license_file.py
import os


def create_license_file():
    license_text = """BSD 3-Clause License

Copyright (c) 2024, XPS Analysis Tool Contributors

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

* Redistributions of source code must retain the above copyright notice, this
  list of conditions and the following disclaimer.

* Redistributions in binary form must reproduce the above copyright notice,
  this list of conditions and the following disclaimer in the documentation
  and/or other materials provided with the distribution.

* Neither the name of the copyright holder nor the names of its
  contributors may be used to endorse or promote products derived from
  this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE."""

    try:
        # Get current working directory
        current_dir = os.getcwd()
        license_path = os.path.join(current_dir, 'LICENSE')

        print(f"Creating LICENSE file at: {license_path}")

        with open(license_path, 'w', encoding='utf-8') as f:
            f.write(license_text)

        print("LICENSE file created successfully!")

        # Verify the file was created
        if os.path.exists(license_path):
            print(f"File confirmed at: {license_path}")
            print(f"File size: {os.path.getsize(license_path)} bytes")
        else:
            print("Error: File was not created")

    except Exception as e:
        print(f"Error creating LICENSE file: {e}")
        print(f"Current directory: {os.getcwd()}")
        print(f"Directory contents: {os.listdir('.')}")


if __name__ == "__main__":
    create_license_file()
    input("Press Enter to continue...")  # Keep window open to see output