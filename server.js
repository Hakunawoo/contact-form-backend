const express = require("express");
const nodemailer = require("nodemailer");
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
    "https://outlook.office365.com/.default"
  );
  return token.token;
}

app.post("/contact", async (req, res) => {
  const { name, email, message } = req.body;

  try {
    const accessToken = await getAccessToken();

    let transporter = nodemailer.createTransport({
      host: "smtp.office365.com",
      port: 587,
      secure: false,
      auth: {
        type: "OAuth2",
        user: userEmail,
        accessToken: accessToken,
      }
    });

    await transporter.sendMail({
      from: `"${name}" <contact@devapps-group.tech>`,
      to: userEmail,
      subject: "New Contact Form Message",
      text: `Name: ${name}\nEmail: ${email}\nMessage:\n${message}`
    });

    res.sendStatus(200);
  } catch (error) {
    console.error(error);
    res.sendStatus(500);
  }
});

app.listen(3000, () => {
  console.log("Server running on http://localhost:3000");
});
