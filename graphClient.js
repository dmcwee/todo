const { Client } = require('@microsoft/microsoft-graph-client');

class GraphClient {
    constructor(token) {
        console.log(`[GraphClient::constructor]`);
        if(!token || !token.trim()) {
            console.log(`[GraphClient::constructor]: Invalid Token`);
            throw new Error('GraphClient: Invalid token received.');
        }

        this._token = token;

        this.graphClient = Client.init({
            authProvider:(done) => {
                done(null, this._token);
            }
        });
    }

    async getTaskList() {
        return await this.graphClient.api('https://graph.microsoft.com/beta/me/todo/lists').get();
    }

    async getDefaultTasks() {
        return await this.graphClient.api('https://graph.microsoft.com/beta/me/todo/lists/Tasks/tasks').get();
    }

    async createNewTask(task) {
        return await this.graphClient.api('https://graph.microsoft.com/beta/me/todo/lists/Tasks/tasks').post(task);
    }
}

exports.GraphClient = GraphClient;