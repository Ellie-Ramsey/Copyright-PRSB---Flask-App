// MANUAL SUBMISSION FORM CODE
$(document).ready(function(){
    $('#myForm').on('submit', function(event){
        event.preventDefault();
        
        // Disable the submit button and show the loading spinner
        $('#submitButton').prop('disabled', true);
        $('#loadingSpinner').show();
        
        $.ajax({
            url: '/',
            method: 'POST',
            data: $(this).serialize(),
            success: function(response){
                $('#resultMessage').html('<div class="alert alert-success">' + response.message + '</div>');
                // Enable the submit button and hide the loading spinner
                $('#submitButton').prop('disabled', false);
                $('#loadingSpinner').hide();
            },
            error: function(){
                $('#resultMessage').html('<div class="alert alert-danger">There was an error submitting the form.</div>');
                // Enable the submit button and hide the loading spinner
                $('#submitButton').prop('disabled', false);
                $('#loadingSpinner').hide();
            }
        });
    });

    // TOGGLE SWITCH CODE
    const toggleSwitch = document.getElementById('toggleSwitch');
    const statusText = document.getElementById('statusText');

    toggleSwitch.addEventListener('change', function() {
        if (toggleSwitch.checked) {
            statusText.textContent = 'Active';
            $.ajax({
                url: '/resume_automation',
                method: 'POST',
                success: function(response){
                    console.log(response.message);
                },
                error: function(){
                    console.error('Failed to resume automation.');
                }
            });
        } else {
            statusText.textContent = 'Paused';
            $.ajax({
                url: '/pause_automation',
                method: 'POST',
                success: function(response){
                    console.log(response.message);
                },
                error: function(){
                    console.error('Failed to pause automation.');
                }
            });
        }
    });
});
