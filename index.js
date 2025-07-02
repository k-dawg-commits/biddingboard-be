const xlsx = require("xlsx")
const express = require("express")
const app = express()
const cors = require("cors")
const authorize = require("./services/googleApiAuthService.js")
const getUnreadMails = require("./services/mailListener")

const PORT = 3222

const filename = "./bid_sheets_merged.xlsx";
const workbook = xlsx.readFile(filename)
const sheetName = "bid_sheets_merged_donttouchpls";
const sheet = workbook.Sheets[sheetName]

let data = xlsx.utils.sheet_to_json(sheet).map((entry, index) => {
    const { __EMPTY, "Bid Count": _, ...newEntry } = entry;

    const excelSerial = entry.Date;
    let formattedDate = null;

    if (typeof excelSerial === "number") {
        const jsDate = new Date((excelSerial - 25569) * 86400 * 1000);
        formattedDate = jsDate.toISOString().split("T")[0];
    } else if (typeof excelSerial === "string") {
        const jsDate = new Date(excelSerial);
        formattedDate = !isNaN(jsDate) ? jsDate.toISOString().split("T")[0] : null;
    }

    return {
        Id: index + 1,
        ...newEntry,
        Date: formattedDate
    };
});


app.use(express.json())
app.use(cors())

const headers = xlsx.utils.sheet_to_json(sheet, { header: 1 })[0].filter(header => header != null)

app.get("/columns-headers", (req, res) => {
    res.json(headers)
})

app.get("/march-bid-sheet", (req, res) => {
    const page = req.query.page || 1
    const pageSize = parseInt(req.query.pageSize) || 10

    const start = (page - 1) * pageSize
    const end = start + pageSize
    const paginatedData = data.slice(start, end)

    res.json({
        page,
        pageSize,
        total: data.length,
        totalPages: Math.ceil(data.length / pageSize),
        entries: paginatedData
    });
})

app.get("/march-bid-sheet-filtered", (req, res) => {
    const page = parseInt(req.query.page) || 1
    const pageSize = parseInt(req.query.pageSize) || 10
    const filter = req.query.filter
    const customerFilter = req.query.customer
    const equipmentFilter = req.query.equipment
    const pickFilter = req.query.pick
    const delFilter = req.query.del
    const fromDate = req.query.fromDate
    const toDate = req.query.toDate

    console.log("Filters:", { filter, customerFilter, equipmentFilter, pickFilter, delFilter, fromDate, toDate })

    let filtered = data

    // Status filter (WON/LOST/TBD)
    if (filter && ["won", "lost", "tbd"].includes(filter.toLowerCase())) {
        filtered = filtered.filter(bid =>
            bid["Won/Lost"] && bid["Won/Lost"].toLowerCase() === filter.toLowerCase()
        )
    }

    // Customer filter
    if (customerFilter && customerFilter !== "all") {
        filtered = filtered.filter(bid =>
            bid.Customer && bid.Customer.toLowerCase().includes(customerFilter.toLowerCase())
        )
    }

    // Equipment filter
    if (equipmentFilter && equipmentFilter !== "all") {
        filtered = filtered.filter(bid =>
            bid.Equipment && bid.Equipment.toLowerCase().includes(equipmentFilter.toLowerCase())
        )
    }

    // Pick filter - handle city, state format with flexible matching
    if (pickFilter && pickFilter !== "all") {
        filtered = filtered.filter(bid => {
            if (!bid.Pick) return false;

            const bidPick = bid.Pick.toLowerCase();
            const filterLower = pickFilter.toLowerCase().trim();

            // Direct match
            if (bidPick.includes(filterLower)) return true;

            // Parse city, state format from filter (e.g., "chicago, il")
            const filterParts = filterLower.split(',');
            if (filterParts.length === 2) {
                const filterCity = filterParts[0].trim();
                const filterState = filterParts[1].trim().toUpperCase(); // Convert state to uppercase for matching

                // Check if bid location contains the city and state (case insensitive)
                const bidPickUpper = bid.Pick.toUpperCase();
                return bidPick.includes(filterCity) && bidPickUpper.includes(filterState);
            }

            return false;
        })
    }

    // Del filter - handle city, state format with flexible matching
    if (delFilter && delFilter !== "all") {
        filtered = filtered.filter(bid => {
            if (!bid.Del) return false;

            const bidDel = bid.Del.toLowerCase();
            const filterLower = delFilter.toLowerCase().trim();

            // Direct match
            if (bidDel.includes(filterLower)) return true;

            // Parse city, state format from filter (e.g., "chicago, il")
            const filterParts = filterLower.split(',');
            if (filterParts.length === 2) {
                const filterCity = filterParts[0].trim();
                const filterState = filterParts[1].trim().toUpperCase(); // Convert state to uppercase for matching

                // Check if bid location contains the city and state (case insensitive)
                const bidDelUpper = bid.Del.toUpperCase();
                return bidDel.includes(filterCity) && bidDelUpper.includes(filterState);
            }

            return false;
        })
    }

    // Date range filter - only apply when both dates are provided
    if (fromDate && toDate) {
        filtered = filtered.filter(bid => {
            if (!bid.Date) return false
            return bid.Date >= fromDate && bid.Date <= toDate
        })
    }

    const total = filtered.length
    const totalPages = Math.ceil(total / pageSize)
    const start = (page - 1) * pageSize
    const end = start + pageSize
    const entries = filtered.slice(start, end)

    res.json({
        page,
        pageSize,
        total,
        totalPages,
        entries
    })
})

app.get("/march-bid-sheet/lost", (req, res) => {
    const page = parseInt(req.query.page) || 1
    const pageSize = parseInt(req.query.pageSize) || 10

    const filtered = data.filter(bid => bid["Won/Lost"] && bid["Won/Lost"].toLowerCase() === "lost")

    const total = filtered.length
    const totalPages = Math.ceil(total / pageSize)
    const start = (page - 1) * pageSize
    const end = start + pageSize
    const entries = filtered.slice(start, end)

    res.json({
        page,
        pageSize,
        total,
        totalPages,
        entries
    })
})

app.get("/march-bid-sheet/tbd", (req, res) => {
    const page = parseInt(req.query.page) || 1
    const pageSize = parseInt(req.query.pageSize) || 10

    const filtered = data.filter(bid => bid["Won/Lost"] && bid["Won/Lost"].toLowerCase() === "tbd")

    const total = filtered.length
    const totalPages = Math.ceil(total / pageSize)
    const start = (page - 1) * pageSize
    const end = start + pageSize
    const entries = filtered.slice(start, end)

    res.json({
        page,
        pageSize,
        total,
        totalPages,
        entries
    })
})

app.post("/add-new-entry", (req, res) => {
    try {
        const newEntry = req.body

        data.push({
            Id: data.length + 1,
            ...newEntry
        })

        const newSheet = xlsx.utils.json_to_sheet(data.map(entry => {
            const { Id, ...rest } = entry
            return rest
        }))

        workbook.Sheets[sheetName] = newSheet
        xlsx.writeFile(workbook, filename)

        res.json({ success: true, message: "Entry added successfully" })
    } catch (error) {
        console.error("Error adding entry:", error)
        res.status(500).json({ success: false, message: "Error adding entry" })
    }
})

app.get("/getzipcodeepls", async (req, res) => {
    const { zipCode } = req.query

    try {
        const response = await fetch(`http://api.zippopotam.us/us/${zipCode}`)
        if (!response.ok) {
            return res.status(404).json({ error: "ZIP code not found" })
        }
        const data = await response.json()
        res.json(data)
    } catch (error) {
        console.error("Error fetching ZIP code data:", error)
        res.status(500).json({ error: "Error fetching ZIP code data" })
    }
})

app.post("/get-history", (req, res) => {
    const { Pick, Del } = req.body

    const history = data.filter(bid =>
        bid.Pick && bid.Del &&
        bid.Pick.toLowerCase().trim() === Pick.toLowerCase().trim() &&
        bid.Del.toLowerCase().trim() === Del.toLowerCase().trim()
    ).sort((a, b) => new Date(b.Date) - new Date(a.Date))

    res.json(history)
})

app.patch("/patch-bid-with-new-status", (req, res) => {
    const { id, newStatus } = req.body

    data[parseInt(id)]["Won/Lost"] = newStatus

    const newSheet = xlsx.utils.json_to_sheet(data.map(entry => {
        const { Id, ...rest } = entry
        return rest
    }))

    workbook.Sheets[sheetName] = newSheet
    xlsx.writeFile(workbook, filename)

    res.json({ success: true, message: "Status updated"})
})

app.get("/num-of-bids", (req, res) => {
    const all = data.length
    const won = data.filter(bid => bid["Won/Lost"] === "WON").length
    const lost = data.filter(bid => bid["Won/Lost"] === "LOST").length
    const tbd =  data.filter(bid => bid["Won/Lost"] === "TBD").length
    const currentDate = new Date().toISOString().split("T")[0]
    const dateOfFirstDayOfCurrentMonth = currentDate.substring(0, 7) + "-01"

    const thisMonth = data.filter(bid => {
        if (!bid.Date) return false
        return bid.Date >= dateOfFirstDayOfCurrentMonth && bid.Date <= currentDate
    }).length


    res.status(200).json({
        all,
        won,
        lost,
        tbd,
        thisMonth
    });
})

app.get("/graph-data", (req, res) => {
    const dateGroups = {};

    data.forEach(bid => {
        const date = bid.Date || 'Unknown';
        if (!dateGroups[date]) {
            dateGroups[date] = { date, bids: 0, won: 0, lost: 0, tbd: 0 };
        }

        dateGroups[date].bids++;

        const status = bid["Won/Lost"];
        if (status === "WON") {
            dateGroups[date].won++;
        } else if (status === "LOST") {
            dateGroups[date].lost++;
        } else if (status === "TBD") {
            dateGroups[date].tbd++;
        }
    });

    const graphData = Object.values(dateGroups)
        .filter(entry => entry.date !== 'Unknown')
        .sort((a, b) => new Date(a.date) - new Date(b.date));

    res.json(graphData);
})

app.get("/revenue-by-customers", (req, res) => {
    const revenueByCustomers = Object.groupBy(data.filter(bid => bid.Customer !== "CHOOSE A CUSTOMER" && bid.Customer.length > 0), bid => bid.Customer)
    const customers = Object.keys(revenueByCustomers)
    const tableData = customers.map(customer => {
        const customerData = revenueByCustomers[customer]
        return {
            Name: customer,
            Bids: customerData.length,
            Won: customerData.filter(bid => bid["Won/Lost"] === "WON").length,
            Lost: customerData.filter(bid => bid["Won/Lost"] === "LOST").length,
            Revenue: customerData
                .filter(bid => bid["Won/Lost"] === "WON")
                .reduce((acc, curr) => {
                    const rawRate = curr[" Bid Rate "]
                    const rate = parseInt(rawRate)

                    if (!isNaN(rate) && rate > 0) {
                        return acc + rate;
                    } else {
                        return acc;
                    }
                }, 0)
        }
    })

    res.status(200).json(tableData)
})


app.get("/recent-bids", (req, res) => {
    res.status(200).json(data.slice(-5))
})

app.get("/mailed-bid-requests", async (req, res) => {
    const auth = await authorize().then().catch(console.error)
    const values = await getUnreadMails(auth).catch(console.error)
    res.status(200).json(values)
})


app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`)
})