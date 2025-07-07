const express = require("express");
const axios = require("axios");
const { ClientSecretCredential } = require("@azure/identity");
require("dotenv").config();

const app = express();
app.use(express.json());

const tenantId = process.env.TENANT_ID;
const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const userEmail = process.env.USER_EMAIL;

const credential = new ClientSecretCredential(
  tenantId,
  clientId,
  clientSecret
);

async function getAccessToken() {
  const token = await credential.getToken(
    "https://graph.microsoft.com/.default"
  );
  return token.token;
}

async function sendMailViaGraph(name, senderEmail, message, token) {
  const emailBody = `
Name: ${name}
Email: ${senderEmail}
Message:
${message}
`;

  const payload = {
    message: {
      subject: "New Contact Form Message",
      body: {
        contentType: "Text",
        content: emailBody
      },
      toRecipients: [
        {
          emailAddress: {
            address: userEmail
          }
        }
      ]
    },
    saveToSentItems: "true"
  };

  await axios.post(
    `https://graph.microsoft.com/v1.0/users/${userEmail}/sendMail`,
    payload,
    {
      headers: {
        Authorization:  `Bearer ${token}`,
        "Content-Type": "application/json"
      }
    }
  );
}

app.post("/contact", async (req, res) => {
  const { name, email, message } = req.body;

  try {
    const token = await getAccessToken();

    await sendMailViaGraph(name, email, message, token);

    res.sendStatus(200);
  } catch (error) {
    console.error("Graph API error:", error.response?.data || error.message);
    res.status(500).send("Error sending email via Graph API");
  }
});

app.get("/test-email", async (req, res) => {
  try {
    const token = await getAccessToken();

    // Send a simple test message
    const payload = {
      message: {
        subject: "TEST EMAIL from Graph API",
        body: {
          contentType: "Text",
          content: "This is a test from the DEVAPPS backend using Microsoft Graph API!"
        },
        from: {
          emailAddress: {
            address: userEmail,
            name: "DEVAPPS GROUP"
          }
        },
        toRecipients: [
          {
            emailAddress: {
              address: userEmail
            }
          }
        ]
      },
      saveToSentItems: "true"
    };

    await axios.post(
      `https://graph.microsoft.com/v1.0/users/${userEmail}/sendMail`,
      payload,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json"
        }
      }
    );

    console.log("Test email sent via Graph API!");
    res.send("Test email sent via Graph API!");
  } catch (error) {
    console.error("Graph API error:", error.response?.data || error.message);
    res.status(500).send("Error sending test email via Graph API");
  }
});

// IMPORTANT: Use Render-assigned PORT if available
const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
