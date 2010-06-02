#include <iostream>
#include <fstream>
#include <string>
#include <tchar.h>
#include <stdio.h>
#include <time.h>
#include <ctime>
#include <cstdlib>
#include <sstream>
#include <vector>

using namespace std;

int main(int argc, _TCHAR* argv[])
{
    char _header[3];
    char _version;
    char _filename[128];
    const int _offset = 0x08;

    cout << "VTF Version Changer v1.0 by Agret <alias.zero2097@gmail.com>" << endl;

    if (argc == 1)
    {
        cout << "Usage: " << argv[0] << " [path/filename]" << endl << "Example: " << argv[0] << " *.vtf" << endl;
        system("pause");
    }
    else
    {
      for (int i=1;i<argc;i++)
      {
        strcpy(_filename, argv[i]);

        fstream ofp;
        ofp.open(_filename, ios_base::in | ios_base::out | ios::binary);

        if (!ofp.is_open())
        {
            cout << "[X] File \"" << _filename << "\" not found." << endl;
        }
        else
        {
          // Check "VTF" is the first 3 bytes of the file to ensure we have a VTF file.
          ofp.read(_header, 3);
          ostringstream os2;
          os2 << _header;
          string s2 = os2.str();
          s2.resize(3);

          if(s2 == "VTF")
          {
            // Check the version by reading 1 byte from offset 0x08
            ofp.seekg(_offset, ios::beg);
            ofp.read(&_version, 1);

            if (_version == 0x00 || _version == 0x01 || _version == 0x02 || _version == 0x03)
            {
                cout << "[-] " << _filename << "\t";
                // Valid version, lets print which one it is.
                if (_version == 0x00) cout << " - v1.0"; // Does this even exist?
                if (_version == 0x01) cout << " - v1.1";
                if (_version == 0x02) cout << " - v1.2";
                if (_version == 0x03) cout << " - v1.3";
                cout << " - no need to patch." << endl;
            }else{
                cout << "[*] " << _filename << "\t";
                // Not a valid version, lets patch the file so it thinks it is =]
                if (_version == 0x04){
                    cout << " - v1.4";
                }else{
                    cout << " - unknown version";
                }

                ofp.seekp(_offset, ios::beg);
                ofp.put(0x03);
                ofp.close();
                cout << " - patched byte." << endl;
            }
          }else{
            cout << "[X] File \"" << _filename << "\" is not a VTF file." << endl;
          }
        }
      }
    }
}
