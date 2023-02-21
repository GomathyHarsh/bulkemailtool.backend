require('dotenv').config();
const express = require('express');
const cors = require("cors");
const db=require('./db/connection');
const jwt = require('jsonwebtoken');


const cookieParser = require('cookie-parser');
const app=express();


//Importing routes
const authRoutes=require('./routes/auth.routes');

const userRoutes=require('./routes/user.routes');
const uploadRoutes = require("./upload");


//Connecting DB
db();

app.use(express.json());
app.use(cookieParser());
app.use(cors({
   origin: 'http://localhost:3000',
   credentials:true
}
));
app.use(express.urlencoded({ extended: true }));

app.get('/',(req,res)=>{
    res.send('Welcome To Email Sending Tool');
})  

app.use('/api',authRoutes);

app.use('/api',userRoutes);
app.use("/api", uploadRoutes);


const PORT = process.env.PORT || 5000;  

app.listen(PORT, () => {
    console.log(`App is running on PORT ${PORT}`);
})