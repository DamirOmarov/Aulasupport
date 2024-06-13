const mongoose = require('mongoose');

const connectDB = async () => {
    try {
        await mongoose.connect(process.env.MONGODB_URI || 'mongodb://aulasupport:B5wGOvwgemfR@10.10.100.186:27017/aula?authSource=admin', {
            useNewUrlParser: true,
            useUnifiedTopology: true,
            serverSelectionTimeoutMS: 5000, // Таймаут после 5 секунд вместо 30
            connectTimeoutMS: 10000, // Таймаут подключения после 10 секунд
        });
        console.log('Connected to MongoDB');
    } catch (err) {
        console.error('Could not connect to MongoDB...', err);
        process.exit(1); // Останавливает приложение при ошибке подключения
    }
};

module.exports = connectDB;
