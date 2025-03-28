while ($true) {
    try {
        # Execute the main.py script
        Write-Host "Running main.py..."
        python main.py

        # Wait for a space to be pressed
        Write-Host "Press space to exit..."

        # Read a single character without displaying it
        $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | Out-Null
        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown").Character

        # Check if the pressed key is a space
        if ($key -eq " ") {
            Write-Host "Exiting..."
            exit 0
        } else {
            Write-Host "Not a space key. Restarting main.py..."
            Start-Sleep -Seconds 1
        }
    }
    catch {
        Write-Error "An error occurred: $($_.Exception.Message)"
        Write-Host "Restarting in 5 seconds..."
        Start-Sleep -Seconds 5
    }
}