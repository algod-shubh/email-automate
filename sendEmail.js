const fs = require('fs');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const { exit } = require('process');


const workbook = xlsx.readFile('./recruiters.xlsx'); 
const sheetName = 'Sheet2';
const worksheet = workbook.Sheets[sheetName];
const data = xlsx.utils.sheet_to_json(worksheet);

console.log("data in sheet is::", data);

const newTransporter = ()=>{
  return nodemailer.createTransport({
    pool: true,
    host: "smtp.gmail.com",
    port: 465,
    secure: true,
    auth: {
      user: "shubhamarora.devs@gmail.com",
      pass: "buii gtob kvla kjzo",      
    },
  });
}
const transporter = newTransporter();
const sendEmail = async (row) => {
  const { Name, Company, Email, Role = "Software Developer", Link = undefined } = row; 
  const nameParts = Name.split(' ');
  const name = nameParts[0];
  const mailOptions = {
    from: 'Shubham Arora <shubhamarora.devs@gmail.com>',
    to: Email,
    subject: `Request for an Interview Opportunity - ${Role} at ${Company}`,
    html: `
    <p>Dear ${name},</p>

    <p>I hope this message finds you well. My name is Shubham Arora, and I am a Full Stack Software Engineer at Paisabazaar.com. I came across your LinkedIn post mentioning that <b>${Company}</b> is looking for a <b>${Role}</b>. I am reaching out to express my interest in the position and provide you with an overview of my qualifications and experience.</p>
    
    <p>I bring a wealth of experience and technical expertise, including:</p>
    
    <ul>
      <li><b>Front-End Development</b>: Expertise in React.js, Next.js, Vue.js, Material-UI, Tailwind CSS, and HTML5/CSS3.</li>
      <li><b>Back-End Development</b>: Proficiency in Node.js, REST APIs, GraphQL, PostgreSQL, MySQL, and Redis.</li>
      <li><b>DevOps & Tools</b>: Hands-on experience with Docker, Jenkins, Firebase, Babel, Webpack, and AWS.</li>
      <li><b>Software Design</b>: Strong knowledge of data structures, algorithms, and scalable application architecture.</li>
    </ul>
    
    <p>Over the years, I have worked on impactful projects, including:</p>
    
    <ul>
      <li>Developing an enterprise-grade data access platform that optimized data querying and improved security compliance.</li>
      <li>Enhancing website performance by improving page load speeds and implementing Progressive Web App (PWA) features.</li>
      <li>Building scalable, modular applications with reusable components, improving development efficiency and user experience.</li>
    </ul>
    
    <p>If you find me suitable, I would greatly appreciate your help with an Interview Opportunity at <b>${Company}</b>.</p>
    
    <p>For your reference, I have attached my <b><a href="https://drive.google.com/file/d/1OfZiqAxEaSozMNl7EHfHCEgRwJ41kU26/view?usp=sharing">Resume</a></b> and my <b><a href="https://www.linkedin.com/in/shubham-arora-dev/">LinkedIn Profile</a></b>. ${Link !== undefined ? `I have also included the relevant <b><a href=${Link}>${Role}</a> opening</b>` : ''} for your convenience.</p>
    
    <p>Looking forward to the opportunity to discuss how I can contribute to the team at <b>${Company}</b>. Thank you for considering my application.</p>
    
    <p>Best regards,</p>
    <p>
    <b>Shubham Arora</b><br>
    Full Stack Software Engineer<br>
    <b>Email:</b> shubhamarora.devs@gmail.com<br>
    <b>Phone:</b> +91 9999014725<br>
    </p>`
    
  };

  try {
    await transporter.sendMail(mailOptions);
    console.log('Email sent to', Email);
  } catch (error) {
    console.error('Error sending email:', Email, error);
  }
};

const sendEmailsSynchronously = async () => {
  for (const row of data) {
    await sendEmail(row);
    await new Promise((resolve) => setTimeout(resolve, Math.random()*90000)); 
  }
  console.log("Done Sending mails")
  exit()
};

sendEmailsSynchronously(); 