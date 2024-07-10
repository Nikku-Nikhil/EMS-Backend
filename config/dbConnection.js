const mongoose = require("mongoose");

const connectDb = async () => {
  try {
    const connect = await mongoose.connect(process.env.CONNECTION_STRING);
    console.log(connect.connection.host, connect.connection.name);
    console.log("Connected to MongoDB");
  } catch (error) {
    console.log(error);
    console.error("Could not connect to MongoDB", error);
    process.exit(1);
  }
};

module.exports = connectDb;
