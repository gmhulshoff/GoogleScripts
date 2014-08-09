function createPainMessage(debtors, event) {
   this.getDrctDbtTxInfXml = function (debtor) {
      return "<DrctDbtTxInf xmlns='" + nameSpace + "'>" +
				"<PmtId><EndToEndId>" + event.parameter.endtoendid + "-" + debtor.mndtid + "</EndToEndId></PmtId>" +
				"<InstdAmt Ccy='EUR'>" + debtor.instdamt + "</InstdAmt>" + 
				"<DrctDbtTx>" + 
					"<MndtRltdInf>" +
						"<MndtId>" + debtor.mndtid + "</MndtId>" +
						"<DtOfSgntr>" + debtor.dtofsgntr + "</DtOfSgntr>" +
						"<AmdmntInd>" + debtor.amdmntind + "</AmdmntInd>" +
					"</MndtRltdInf>" + 
				"</DrctDbtTx>" +
				"<DbtrAgt><FinInstnId/></DbtrAgt>" + 
				"<Dbtr>" +
					"<Nm>" + debtor.dbtrnm + "</Nm>" +
					"<PstlAdr><Ctry>NL</Ctry></PstlAdr>" + 
				"</Dbtr>" +
				"<DbtrAcct><Id><IBAN>" + debtor.dbtriban + "</IBAN></Id></DbtrAcct>" +
				"<UltmtDbtr><Nm>" + debtor.dbtrnm + "</Nm></UltmtDbtr>" +
				"<RmtInf><Ustrd>" + debtor.dbtrustrd + "</Ustrd></RmtInf>" + 
			"</DrctDbtTxInf>";
   }
   this.getPmtInfXml = function (paymentType) {
      return "<PmtInf xmlns='" + nameSpace + "'>" +
			"<PmtInfId>" + event.parameter.endtoendid + '-' + paymentType + "</PmtInfId>" +
			"<PmtMtd>DD</PmtMtd>" +
			"<PmtTpInf>" +
				"<SvcLvl><Cd>SEPA</Cd></SvcLvl>" +
				"<LclInstrm><Cd>CORE</Cd></LclInstrm>" +
				"<SeqTp>" + paymentType + "</SeqTp>" +
			"</PmtTpInf>" +
			"<ReqdColltnDt>" + event.parameter.incassodate.split('T')[0] + "</ReqdColltnDt>" +
			"<Cdtr>" +
				"<Nm>" + event.parameter.initiatingparty + "</Nm>" +
				"<PstlAdr><Ctry>NL</Ctry></PstlAdr>" +
			"</Cdtr>" +
			"<CdtrAcct><Id><IBAN>" + event.parameter.initiatingiban + "</IBAN></Id></CdtrAcct>" +
			"<CdtrAgt><FinInstnId><BIC>" + event.parameter.initiatingbic + "</BIC></FinInstnId></CdtrAgt>" +
			"<UltmtCdtr><Nm>" + event.parameter.initiatingparty + "</Nm></UltmtCdtr>" +
			"<ChrgBr>SLEV</ChrgBr>" +
			"<CdtrSchmeId>" +
				"<Id>" +
					"<PrvtId>" +
						"<Othr>" +
							"<Id>" + event.parameter.initiatingcreditoridentifier + "</Id>" +
							"<SchmeNm><Prtry>SEPA</Prtry></SchmeNm>" +
						"</Othr>" +
					"</PrvtId>" +
				"</Id>" +
			"</CdtrSchmeId>" +
		"</PmtInf>";
   }
   this.getDocumentXml = function () {
      return "<Document xmlns='" + nameSpace + "' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>" +
			"<CstmrDrctDbtInitn>" +
				"<GrpHdr>" +
					"<MsgId>" + event.parameter.endtoendid + event.parameter.creationdatetime.split('T')[0] + "</MsgId>" +
					"<CreDtTm>" + event.parameter.creationdatetime + "</CreDtTm>" +
					"<NbOfTxs>" + Utilities.formatString('%d', nbOfTxsFRST + nbOfTxsRCUR) + "</NbOfTxs>" +
					"<InitgPty><Nm>" + event.parameter.initiatingparty + "</Nm></InitgPty>" +
				"</GrpHdr>" +
			"</CstmrDrctDbtInitn>" +
		  "</Document>";
   }

   this.firstFlag = 'FRST';
   this.recurFlag = 'RCUR';
   this.nameSpace = "urn:iso:std:iso:20022:tech:xsd:pain.008.001.02";
   this.ns = XmlService.getNamespace(nameSpace);
   this.PmtInfFRST = XmlService.parse(getPmtInfXml(firstFlag)).detachRootElement();
   this.PmtInfRCUR = XmlService.parse(getPmtInfXml(recurFlag)).detachRootElement();
   this.nbOfTxsFRST = 0; 
   this.nbOfTxsRCUR = 0;
   this.updateTransactionCount = function(debtor) { if (debtor["seqtp"] == firstFlag) nbOfTxsFRST++; else nbOfTxsRCUR++; }
   this.addTransaction = function(pmtInf) { 
      pmtInf.addContent(XmlService.parse(getDrctDbtTxInfXml(debtor)).detachRootElement()); 
      updateTransactionCount(debtor);
   }
   
   for (var i = 0, debtor; debtor = debtors[i]; i++)
     if (debtor.dbtriban != '')
       addTransaction(debtor["seqtp"] == firstFlag ? PmtInfFRST : PmtInfRCUR);
   
   var rootDocument = XmlService.parse(getDocumentXml());
   var cstmrDrctDbtInitnXml = rootDocument.getRootElement().getChild("CstmrDrctDbtInitn", ns);
   if (nbOfTxsFRST > 0)
	  cstmrDrctDbtInitnXml.addContent(PmtInfFRST);
   if (nbOfTxsRCUR > 0)
	  cstmrDrctDbtInitnXml.addContent(PmtInfRCUR);
   
   return XmlService.getPrettyFormat().setLineSeparator('\n').format(rootDocument);
}
