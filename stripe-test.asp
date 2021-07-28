<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/sql_connection.asp" -->
<!--#include virtual="/Connections/chilkat.asp" -->
<!--#include virtual="/Connections/stripe.asp" -->
<!DOCTYPE html>
<html>
    <head>
        <title>Buy cool new product</title>
        <script src="https://polyfill.io/v3/polyfill.min.js?version=3.52.1&features=fetch"></script>
        <script src="https://js.stripe.com/v3/"></script>
      </head>
<body>

        <section>
          <button type="button" id="checkout-button">Checkout</button> Session ID: <%= stripe_checkout_session_id %>
        </section>

      <script type="text/javascript">
        // Create an instance of the Stripe object with your publishable API key
        var stripe = Stripe("<%= STRIPE_PUBLISHABLE_KEY %>");
        var checkoutButton = document.getElementById("checkout-button");
    
        checkoutButton.addEventListener("click", function () {
          fetch("/stripe/create-checkout-session.asp", {
            method: "POST",
          })
            .then(function (response) {
              return response.json();
            })
            .then(function (session) {
              return stripe.redirectToCheckout({ sessionId: session.id });
            })
            .then(function (result) {
              // If redirectToCheckout fails due to a browser or network
              // error, you should display the localized error message to your
              // customer using error.message.
              if (result.error) {
                alert(result.error.message);
              }
            })
            .catch(function (error) {
              console.error("Error:", error);
            });
        });





        
      </script>
</body>
</html>