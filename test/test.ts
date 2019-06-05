import * as assert from "assert";
import * as fs from "fs";
import * as mocha from "mocha";
import * as path from "path";
import * as testHelper from "office-addin-test-helpers";
import * as testServerInfra from "office-addin-test-server";
import * as WebSocket from "ws";

const host: string = "excel";
const manifestPath = path.resolve(`${process.cwd()}/test/manifest.xml`);
const port: number = 4201;
const testDataFile: string = `${process.cwd()}/test/src/testData.json`;
const testJsonData = JSON.parse(fs.readFileSync(testDataFile).toString());
const testServer = new testServerInfra.TestServer(port);
let testValues: any = [];
let connected = false;
let events = [];

function initializeDebugger(){
    let ws = new WebSocket('ws://127.0.0.1:9229/runtime1');
    ws.onopen = function () {
        connected = true;
        console.log('socket connection opened');
        ws.send("{\"id\":1,\"method\":\"Console.enable\"}")
        ws.send("{\"id\":2,\"method\":\"Debugger.enable\"}")
        ws.send("{\"id\":3,\"method\":\"Runtime.enable\"}")
        ws.send("{\"id\":5,\"method\":\"Runtime.runIfWaitingForDebugger\"}")
        // ws.send("{\"id\":23,\"method\":\"Debugger.resume\"}")
    };

    ws.onmessage = function (event) {
        const data = JSON.parse(event.data.toString())
        if(data["method"] === "Runtime.consoleAPICalled"){
            console.log(event.data);
            events.push(data["params"]);
        }
    };
    ws.onclose = function(){
        if (!connected){
            setTimeout(function(){initializeDebugger()}, 1000);
        }
        else{
            console.log("Connection closed...");
        }
    };
    ws.onerror = function (event) {
    };
}

describe("Test Excel Custom Functions", function () {
    before("Start test server", async function () {
        const testServerStarted = await testServer.startTestServer(true /* mochaTest */);
        const serverResponse = await testHelper.pingTestServer(port);
        assert.equal(testServerStarted, true);
        assert.equal(serverResponse["status"], 200);
    }),
    describe("Start dev-server and sideload application", function () {
        it(`Sideload should have completed for ${host} and dev-server should have started`, async function () {
            this.timeout(0);
            const startDevServer = await testHelper.startDevServer();
            const sideloadApplication = await testHelper.sideloadDesktopApp(host, manifestPath);
            assert.equal(startDevServer, true);
            assert.equal(sideloadApplication, true);
        });
    });
    describe("Get test results for custom functions and validate results", function () {
        it("should get results from the taskpane application", async function () {
            this.timeout(0);
            // Expecting six result values
            initializeDebugger();
            testValues = await testServer.getTestResults();
            assert.equal(testValues.length, 6);
        });
        it("ADD function should return expected value", async function () {
            assert.equal(testJsonData.functions.ADD.result, testValues[0].Value);
        });
        it("CLOCK function should return expected value", async function () {
            // Check that captured values are different to ensure the function is streaming
            assert.notEqual(testValues[1].Value, testValues[2].Value);
            // Check if the returned string contains 'AM' or 'PM', indicating it's a time-stamp
            assert.equal(true, testValues[1].Value.includes(testJsonData.functions.CLOCK.result.amString) || testValues[1].Value.includes(testJsonData.functions.CLOCK.result.pmString) ? true : false);
            assert.equal(true, testValues[2].Value.includes(testJsonData.functions.CLOCK.result.amString) || testValues[2].Value.includes(testJsonData.functions.CLOCK.result.pmString) ? true : false);
        });
        it("INCREMENT function should return expected value", async function () {
            // Check that captured values are different to ensure the function is streaming
            assert.notEqual(testValues[3].Value, testValues[4].Value);
            // Check to see that both captured streaming values are divisible by 4
            assert.equal(0, testValues[3].Value % testJsonData.functions.INCREMENT.result);
            assert.equal(0, testValues[4].Value % testJsonData.functions.INCREMENT.result);
        });
        it("LOG function should return expected value", async function () {
            assert.equal(testJsonData.functions.LOG.result, testValues[5].Value);
            //protocol validation
            assert.strictEqual(events.length, 1);
            const logEvent = events.shift();
            assert.equal(logEvent["type"], "log");
            assert.strictEqual(logEvent["args"].length, 1);
            assert.equal(logEvent["args"][0]["type"], "string");
            assert.equal(logEvent["args"][0]["description"], "this is a test");
        });
    });
    after("Teardown test environment", async function () {
        const stopTestServer = await testServer.stopTestServer();
        assert.equal(stopTestServer, true);
        const testEnvironmentTornDown = await testHelper.teardownTestEnvironment(host);
        assert.equal(testEnvironmentTornDown, true);
    });
});