for /f "delims=" %%a in ('dir .\examples\*.rs /b/s') do cargo run --example %%~na
