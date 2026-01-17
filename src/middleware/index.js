const resolveGraphToken = require("../auth/tokenSelector");

app.use("/api", resolveGraphToken);
