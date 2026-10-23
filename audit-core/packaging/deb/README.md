# deb packaging gate

This directory is reserved for the eventual self-contained deb layout. Do not
package the current probe executables as a production application. Create the
package after a production CLI exists, real sanitized workbook templates pass
the required preservation checks, and installation plus end-to-end execution
have been accepted on both Kylin V10 and UOS V20.

The current release build and GLIBC symbol scan are run by `../build.sh`.
